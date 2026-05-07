const express = require('express');
const session = require('express-session');
const multer = require('multer');
const sqlite3 = require('sqlite3').verbose();
const path = require('path');
const fs = require('fs');
const mammoth = require('mammoth');
const xlsx = require('xlsx');
const pdfParse = require('pdf-parse');

const app = express();
const port = 3000;

// Setup database
const dbPath = path.join(__dirname, 'database.sqlite');
const db = new sqlite3.Database(dbPath);

db.serialize(() => {
  db.run(`
    CREATE TABLE IF NOT EXISTS materials (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      title TEXT NOT NULL,
      content TEXT,
      media_path TEXT,
      media_type TEXT,
      department TEXT,
      level INTEGER DEFAULT 1,
      parent_id INTEGER DEFAULT NULL,
      created_at DATETIME DEFAULT CURRENT_TIMESTAMP
    )
  `);

  db.run(`CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, value TEXT)`);
  db.run(`CREATE TABLE IF NOT EXISTS document_versions (id INTEGER PRIMARY KEY AUTOINCREMENT, department TEXT, timestamp TEXT, snapshot TEXT, created_at DATETIME DEFAULT CURRENT_TIMESTAMP)`);
  
  db.run(`CREATE TABLE IF NOT EXISTS users (id INTEGER PRIMARY KEY AUTOINCREMENT, username TEXT UNIQUE, password TEXT, department TEXT, department_name TEXT, role TEXT DEFAULT 'editor')`, (err) => {
    if (!err) {
      db.run("INSERT OR IGNORE INTO users (username, password, department, department_name, role) VALUES ('admin', 'admin123', 'admin', '总管理台', 'admin')");
      db.run("INSERT OR IGNORE INTO users (username, password, department, department_name, role) VALUES ('gongcheng', '123456', 'engineering', '工程部', 'editor')");
      db.run("INSERT OR IGNORE INTO users (username, password, department, department_name, role) VALUES ('sheji', '123456', 'design', '设计部', 'editor')");
    }
  });
  
  db.run(`CREATE TABLE IF NOT EXISTS activity_logs (id INTEGER PRIMARY KEY AUTOINCREMENT, username TEXT, department TEXT, action TEXT, details TEXT, created_at DATETIME DEFAULT CURRENT_TIMESTAMP)`);

  // Backward compatibility: Adding columns to old tables if they don't exist
  db.run(`ALTER TABLE materials ADD COLUMN department TEXT DEFAULT 'engineering'`, (err) => { });
  db.run(`ALTER TABLE materials ADD COLUMN level INTEGER DEFAULT 1`, (err) => { });
  db.run(`ALTER TABLE materials ADD COLUMN parent_id INTEGER DEFAULT NULL`, (err) => {
    // Migrate old data to link level 2 items to level 1 items via parent_id
    db.all('SELECT id, level, parent_id FROM materials ORDER BY id ASC', (err, rows) => {
      if (!err && rows) {
        let lastL1 = null;
        rows.forEach(r => {
          if (r.level === 1) lastL1 = r.id;
          else if (r.level === 2 && !r.parent_id && lastL1) {
            db.run('UPDATE materials SET parent_id = ? WHERE id = ?', [lastL1, r.id]);
          }
        });
      }
    });
  });
});

// Configure multer for file uploads
const storage = multer.diskStorage({
  destination: function (req, file, cb) {
    cb(null, 'public/uploads/')
  },
  filename: function (req, file, cb) {
    const uniqueSuffix = Date.now() + '-' + Math.round(Math.random() * 1E9)
    cb(null, uniqueSuffix + path.extname(file.originalname))
  }
});

const upload = multer({ storage: storage });

// App configuration
app.set('view engine', 'ejs');
app.use(express.static('public'));
app.use(express.urlencoded({ extended: true }));
app.use(express.json());
app.use(session({
  secret: 'wendang_secret_key_2026',
  resave: false,
  saveUninitialized: true
}));

// Middleware
function isAuthenticated(req, res, next) {
  if (req.session.loggedIn) return next();
  res.redirect('/login');
}

// 1. Frontend Document Generation
app.get('/', (req, res) => {
  db.all("SELECT DISTINCT department, department_name FROM users WHERE role = 'editor'", (err, depts) => {
    let links = (depts || []).map(d => `<a href="/doc/${d.department}" class="eng-btn" style="background:#3498db; margin-bottom:15px;">📋 ${d.department_name}</a>`).join('');
    res.send(`
    <!DOCTYPE html>
    <html lang="zh-CN">
    <head>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <title>选择文档部门</title>
      <style>
        body { font-family: 'Segoe UI', sans-serif; display: flex; flex-direction: column; align-items: center; justify-content: center; min-height: 100vh; background: #f0f2f5; margin: 0; padding: 20px; }
        .box { background: white; padding: 40px; border-radius: 8px; box-shadow: 0 4px 12px rgba(0,0,0,0.1); text-align: center; width: 100%; max-width: 400px; }
        h1 { color: #2c3e50; margin-top: 0; margin-bottom: 30px; font-size: 1.5rem; }
        a { display: block; margin: 15px auto; padding: 15px 30px; text-decoration: none; border-radius: 4px; font-size: 16px; color: white; transition: background 0.3s; }
        .eng-btn:hover { opacity: 0.9; }
        .admin-link { display: block; margin-top: 30px; color: #7f8c8d; text-decoration: none; background: transparent; padding: 0; }
        .admin-link:hover { color: #34495e; background: transparent; }
      </style>
    </head>
    <body>
      <div class="box">
        <h1>请选择要查看的文档类型</h1>
        ${links}
        <a href="/login" class="admin-link">进入后台管理系统</a>
      </div>
    </body>
    </html>
    `);
  });
});

app.get('/doc/:dept', (req, res) => {
  const dept = req.params.dept;
  
  db.get('SELECT department_name FROM users WHERE department = ? LIMIT 1', [dept], (err, row) => {
    if (!row) return res.status(404).send('未找到该部门文档');
    
    const deptName = row.department_name;

  const versionId = req.query.vid;

  const renderRows = (rows, timestampStr = null) => {
    let level1 = rows.filter(r => r.level === 1);
    let grouped = level1.map(l1 => {
      return { ...l1, children: rows.filter(r => r.level === 2 && r.parent_id === l1.id) };
    });
    let orphans = rows.filter(r => r.level === 2 && !r.parent_id);
    orphans.forEach(o => grouped.push({ ...o, level: 1, children: [] }));

    let displayName = timestampStr ? `${deptName} (历史版本: ${timestampStr})` : deptName;
    res.render('index', { grouped, departmentName: displayName, department: dept, isHistory: !!timestampStr });
  };

  if (versionId) {
    db.get('SELECT snapshot, timestamp FROM document_versions WHERE id = ? AND department = ?', [versionId, dept], (err, row) => {
      if (row && row.snapshot) {
        renderRows(JSON.parse(row.snapshot), row.timestamp);
      } else {
        res.redirect(`/doc/${dept}`);
      }
    });
  } else {
    db.all('SELECT * FROM materials WHERE department = ? ORDER BY id ASC', [dept], (err, rows) => {
      if (err) return res.status(500).send("Database error");
      renderRows(rows);
    });
  }
  }); // End db.get callback
});

// 2. Login Routes
app.get('/login', (req, res) => {
  if (req.session.loggedIn) return res.redirect('/admin');
  res.render('login', { error: null });
});

app.post('/login', (req, res) => {
  const { username, password } = req.body;
  db.get('SELECT * FROM users WHERE username = ? AND password = ?', [username, password], (err, user) => {
    if (user) {
      req.session.loggedIn = true; 
      req.session.username = user.username;
      req.session.department = user.department; 
      req.session.departmentName = user.department_name; 
      res.redirect('/admin');
    } else {
      res.render('login', { error: '账号或密码错误。' });
    }
  });
});

app.get('/logout', (req, res) => {
  req.session.destroy();
  res.redirect('/login');
});

// 3. Admin Routes
app.get('/admin', isAuthenticated, (req, res) => {
  if (req.session.department === 'admin') {
    db.get('SELECT value FROM settings WHERE key = ?', ['deepseek_api_key'], (err, row) => {
      const apiKey = row ? row.value : '';
      db.all('SELECT id, department, timestamp FROM document_versions ORDER BY id DESC', (err, versions) => {
        db.all('SELECT * FROM users WHERE role = "editor"', (err, users) => {
          db.all('SELECT * FROM activity_logs ORDER BY id DESC LIMIT 50', (err, logs) => {
            res.render('admin_super', { apiKey, versions, departmentName: '总管理台', users: users || [], logs: logs || [] });
          });
        });
      });
    });
    return;
  }

  // Normal department users
  db.all('SELECT * FROM materials WHERE department = ? ORDER BY id ASC', [req.session.department], (err, materials) => {
    db.all('SELECT * FROM activity_logs WHERE username = ? ORDER BY id DESC LIMIT 20', [req.session.username], (err, logs) => {
      res.render('admin', { materials, department: req.session.department, departmentName: req.session.departmentName, logs: logs || [] });
    });
  });
});

app.post('/admin/users', isAuthenticated, (req, res) => {
  if (req.session.department !== 'admin') return res.redirect('/admin');
  const { username, password, department, department_name } = req.body;
  db.run("INSERT INTO users (username, password, department, department_name, role) VALUES (?, ?, ?, ?, 'editor')", 
    [username, password, department, department_name], 
    () => res.redirect('/admin')
  );
});

app.post('/admin/users/delete', isAuthenticated, (req, res) => {
  if (req.session.department !== 'admin') return res.redirect('/admin');
  const { id } = req.body;
  db.run("DELETE FROM users WHERE id = ? AND role = 'editor'", [id], () => res.redirect('/admin'));
});

// Settings (Save API Key)
app.post('/admin/settings', isAuthenticated, (req, res) => {
  const { apiKey } = req.body;
  db.run('INSERT OR REPLACE INTO settings (key, value) VALUES (?, ?)', ['deepseek_api_key', apiKey], () => res.redirect('/admin'));
});

// Init outline with level detection
app.post('/admin/init-outline', isAuthenticated, async (req, res) => {
  // admin can initialize outline for any department
  if (req.session.department !== 'admin') return res.redirect('/admin');
  
  const { outline, targetDept } = req.body;
  if (!outline || !targetDept) return res.redirect('/admin');
  if (!outline) return res.redirect('/admin');

  const lines = outline.split('\n').map(l => l.trim()).filter(l => l.length > 0);

  try {
    let currentParentId = null;
    db.serialize(() => {
      db.run('BEGIN TRANSACTION');

      const insertRow = (title, level, parentId) => {
        return new Promise((resolve, reject) => {
          db.run('INSERT INTO materials (title, content, department, level, parent_id) VALUES (?, "", ?, ?, ?)',
            [title, targetDept, level, parentId], function (err) {
              if (err) reject(err); else resolve(this.lastID);
            });
        });
      };

      (async function processLines() {
        for (const title of lines) {
          let level = 2;
          if (/^[一二三四五六七八九十]+、/.test(title)) level = 1;

          let pId = (level === 1) ? null : currentParentId;
          const insertedId = await insertRow(title, level, pId);
          if (level === 1) currentParentId = insertedId;
        }
        db.run('COMMIT', () => res.redirect('/admin'));
      })();
    });
  } catch (e) {
    console.error(e);
    res.redirect('/admin');
  }
});

// AI Generate content & Save to MD
app.post('/admin/generate-ai-content', isAuthenticated, async (req, res) => {
  const { newMaterials } = req.body;
  const dept = req.session.department;

  db.get('SELECT value FROM settings WHERE key = ?', ['deepseek_api_key'], async (err, row) => {
    const apiKey = row ? row.value : null;
    if (!apiKey) return res.send('请先配置 API Key');

    db.all('SELECT id, title, level, parent_id, content FROM materials WHERE department = ? ORDER BY id ASC', [dept], async (err, rows) => {
      if (err || rows.length === 0) return res.redirect('/admin');

      // Pass full content for fusion and deduplication
      const outlineContext = rows.map(r => `ID: ${r.id} | 层级: ${r.level} | 标题: ${r.title} | 现有完整内容: ${r.content || '空'}`).join('\n\n');

      let systemPrompt = `你是一个智能文档助手。负责接收用户的新资料，并更新到现有的文档结构中。\n\n当前大纲结构及各自的【现有完整内容】如下：\n${outlineContext}\n\n`;
      systemPrompt += `规则：\n`;
      systemPrompt += `1. 更新/融合（防止冗余，最高优先级）：如果新资料所讲的知识点在某个现有小节中已经存在相关描述，请你将新资料与该小节的【现有完整内容】进行深度融合、去重和优化，重新撰写出一份结构更清晰、覆盖全面的正文。使用 action: "overwrite" 覆盖该小节。\n`;
      systemPrompt += `2. 追加：仅当新资料属于某小节，且与该小节现有内容是完全独立的不同事项（例如新增的一个无关案例）时，才使用 action: "append" 追加在末尾。\n`;
      systemPrompt += `3. 新建：如果讲述的是完全全新的知识点/大坑，现有大纲中没有合适的小节，选择 action: "create"，指定所属大章节的 parent_id 和新小节的 title。\n`;
      systemPrompt += `4. 极其重要：用户的富文本资料中若包含图片（<img src="...">）或视频（<video src="...">）标签，你必须在你生成的内容中原封不动地保留这些标签，将它们插入到合适的位置！绝不能丢失图片或视频链接！\n`;

      if (dept === 'engineering') {
        systemPrompt += `5. 【工程部专属 - 大章节说明】：请为本次涉及更新的所有大章节（层级为1），生成或补充一段引导性的说明文字（优先尝试 overwrite 融合更新，如果没有旧内容则 append）。\n`;
        systemPrompt += `6. 【工程部专属 - 核心总结】：大纲最后通常有一个“核心总结”节点（比如六、核心总结）。每次处理新资料时，请务必结合本次的新增内容以及整个大纲目前的全面情况，重新撰写并覆盖该总结章节的内容（使用 action: "overwrite" 并指定该总结节点的id）。\n`;
      }

      systemPrompt += `\n请严格返回合法JSON数组。示例格式：\n[\n  {"action": "overwrite", "id": 1, "content": "<p>融合去重后的全量新内容</p>"},\n  {"action": "create", "parent_id": 1, "title": "2. 新坑点", "content": "<p>新正文</p>"}\n]\n只输出JSON！`;

      try {
        const response = await fetch('https://api.deepseek.com/v1/chat/completions', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json', 'Authorization': `Bearer ${apiKey}` },
          body: JSON.stringify({
            model: 'deepseek-v4-pro',
            messages: [
              { role: 'system', content: systemPrompt },
              { role: 'user', content: `富文本新资料：\n${newMaterials}` }
            ],
            temperature: 0.7
          })
        });

        if (!response.ok) return res.status(500).send("API Error: " + await response.text());

        const data = await response.json();
        let text = data.choices[0].message.content.trim();
        if (text.startsWith('\`\`\`json')) text = text.substring(7, text.length - 3).trim();
        else if (text.startsWith('\`\`\`')) text = text.substring(3, text.length - 3).trim();

        const updates = JSON.parse(text);

        // Save generated content to a Markdown file
        const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
        const mdFileName = `doc_${dept}_${timestamp}.md`;
        let mdContent = `# 新增内容归档 - ${new Date().toLocaleString()}\n\n`;
        mdContent += `## 原始输入 (HTML富文本):\n${newMaterials}\n\n`;
        mdContent += `---\n\n## AI 生成归类结果:\n\n`;
        for (const up of updates) {
          if (up.action === 'append') {
            const matchTitle = rows.find(r => r.id == up.id)?.title || '未知小节';
            mdContent += `### [追加] ${matchTitle} (ID: ${up.id})\n${up.content}\n\n`;
          } else if (up.action === 'create') {
            const parentTitle = rows.find(r => r.id == up.parent_id)?.title || '未知大章';
            mdContent += `### [新建小节] ${up.title} (隶属于: ${parentTitle})\n${up.content}\n\n`;
          } else if (up.action === 'overwrite') {
            const matchTitle = rows.find(r => r.id == up.id)?.title || '未知小节';
            mdContent += `### [覆盖重写] ${matchTitle} (ID: ${up.id})\n${up.content}\n\n`;
          }
        }
        fs.writeFileSync(path.join(__dirname, 'public/docs', mdFileName), mdContent);

        // Update DB
        const updateStmt = db.prepare('UPDATE materials SET content = CASE WHEN content IS NULL OR content = "" THEN ? ELSE content || "<br><br>" || ? END WHERE id = ? AND department = ?');
        const overwriteStmt = db.prepare('UPDATE materials SET content = ? WHERE id = ? AND department = ?');
        const createStmt = db.prepare('INSERT INTO materials (title, content, department, level, parent_id) VALUES (?, ?, ?, 2, ?)');

        db.serialize(() => {
          db.run('BEGIN TRANSACTION');
          for (const up of updates) {
            if (up.action === 'append' && up.id && up.content) {
              updateStmt.run(up.content, up.content, up.id, dept);
            } else if (up.action === 'overwrite' && up.id && up.content) {
              overwriteStmt.run(up.content, up.id, dept);
            } else if (up.action === 'create' && up.parent_id && up.title && up.content) {
              createStmt.run(up.title, up.content, dept, up.parent_id);
            } else if (up.id && up.content && !up.action) {
              // fallback for older AI generation format
              updateStmt.run(up.content, up.content, up.id, dept);
            }
          }
          db.run('COMMIT', () => {
            updateStmt.finalize();
            overwriteStmt.finalize();
            createStmt.finalize();

            // Record full snapshot for Version Control
            db.all('SELECT * FROM materials WHERE department = ? ORDER BY id ASC', [dept], (err, currentRows) => {
              if (!err && currentRows) {
                const snapshot = JSON.stringify(currentRows);
                const tsName = new Date().toLocaleString();
                db.run('INSERT INTO document_versions (department, timestamp, snapshot) VALUES (?, ?, ?)', [dept, tsName, snapshot], () => {
                  db.run('INSERT INTO activity_logs (username, department, action, details) VALUES (?, ?, ?, ?)', 
                    [req.session.username, dept, '新增内容 (AI归档)', '新提交了一批内容，并生成了最新快照版本'], () => {
                      res.redirect('/admin'); 
                  });
                });
              } else {
                res.redirect('/admin'); 
              }
            });
          });
        });

      } catch (err) {
        res.status(500).send("AI Processing Error: " + err.message);
      }
    });
  });
});

// Upload media for WangEditor
app.post('/admin/upload-media', isAuthenticated, upload.single('file'), (req, res) => {
  if (!req.file) return res.status(400).json({ errno: 1, message: 'No file' });
  const url = '/uploads/' + req.file.filename;
  // WangEditor v5 expects this exact format
  res.json({
    errno: 0,
    data: {
      url: url,
      alt: '',
      href: ''
    }
  });
});

// Provide manual section update without rich text popup (if needed)
app.post('/admin/add-to-section/:id', isAuthenticated, upload.single('media'), (req, res) => {
  if (!req.file) return res.redirect('/admin');
  const media_path = '/uploads/' + req.file.filename;
  const media_type = req.file.mimetype.startsWith('image/') ? 'image' : (req.file.mimetype.startsWith('video/') ? 'video' : null);
  db.run('UPDATE materials SET media_path = ?, media_type = ? WHERE id = ? AND department = ?', [media_path, media_type, req.params.id, req.session.department], () => res.redirect('/admin'));
});

app.get('/admin/clear-section-content/:id', isAuthenticated, (req, res) => {
  db.run('UPDATE materials SET content = "", media_path = NULL, media_type = NULL WHERE id = ? AND department = ?', [req.params.id, req.session.department], () => res.redirect('/admin'));
});

app.get('/admin/reset-outline', isAuthenticated, (req, res) => {
  db.run('DELETE FROM materials WHERE department = ?', [req.session.department], () => res.redirect('/admin'));
});

// Parse document and extract text
app.post('/admin/parse-document', isAuthenticated, upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file uploaded' });

  const ext = path.extname(req.file.originalname).toLowerCase();
  const filePath = req.file.path;
  let textContent = '';

  try {
    let isHtml = false;

    if (ext === '.txt' || ext === '.md') {
      textContent = fs.readFileSync(filePath, 'utf-8');
    } else if (ext === '.docx') {
      // 使用 convertToHtml 以便在富文本编辑器中保留加粗、列表等基本排版
      const result = await mammoth.convertToHtml({ path: filePath });
      textContent = result.value;
      isHtml = true;
    } else if (ext === '.pdf') {
      const dataBuffer = fs.readFileSync(filePath);
      const data = await pdfParse(dataBuffer);
      textContent = data.text || '';

      // Remove null bytes which might break JSON or rendering
      textContent = textContent.replace(/\0/g, '');

      if (!textContent.trim()) {
        textContent = '<span style="color:red;">（系统提示：未能从该 PDF 中提取出任何文字，这可能是一个纯图片扫描版的 PDF，需要专门的 OCR 软件才能识别文字。）</span>';
        isHtml = true;
      }
    } else if (ext === '.xlsx' || ext === '.xls') {
      const workbook = xlsx.readFile(filePath);
      textContent = '';
      workbook.SheetNames.forEach(sheetName => {
        const sheet = workbook.Sheets[sheetName];
        textContent += `<br/><h3>--- 表格: ${sheetName} ---</h3><br/>`;
        // 使用 sheet_to_html 生成真实的 HTML <table> 标签，完美适配富文本编辑器，解决排版混乱/乱码感
        textContent += xlsx.utils.sheet_to_html(sheet);
      });
      isHtml = true;
    } else {
      return res.status(400).json({ error: 'Unsupported file type. Supported types: .txt, .md, .docx, .pdf, .xlsx, .xls' });
    }

    // 只针对纯文本（PDF, TXT）替换换行符为 <br/>，并包装到 <p> 中，否则可能导致富文本编辑器因为不合规块级元素而渲染失败
    if (!isHtml) {
      textContent = `<p>${(textContent || '').replace(/\n/g, '<br/>')}</p>`;
    }

    res.json({ text: textContent });
  } catch (err) {
    console.error('File parsing error:', err);
    res.status(500).json({ error: 'Failed to parse document: ' + err.message });
  }
});

app.listen(port, () => console.log(`Document server listening at http://localhost:${port}`));
