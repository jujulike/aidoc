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
const aiTasks = {};

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
  db.run(`CREATE TABLE IF NOT EXISTS user_uploads (id INTEGER PRIMARY KEY AUTOINCREMENT, filename TEXT, original_name TEXT, file_path TEXT, file_type TEXT, department TEXT, username TEXT, created_at DATETIME DEFAULT CURRENT_TIMESTAMP)`);

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

  // Migrate existing department outlines into settings if not present
  setTimeout(() => {
    db.all("SELECT DISTINCT department FROM users WHERE role = 'editor'", (err, depts) => {
      if (!err && depts) {
        depts.forEach(d => {
          const dept = d.department;
          db.get("SELECT value FROM settings WHERE key = ?", ['template_' + dept], (err, row) => {
            if (!row) {
              db.all("SELECT title FROM materials WHERE department = ? ORDER BY id ASC", [dept], (err, mRows) => {
                if (mRows && mRows.length > 0) {
                  const templateText = mRows.map(r => r.title).join('\n');
                  db.run("INSERT INTO settings (key, value) VALUES (?, ?)", ['template_' + dept, templateText]);
                }
              });
            }
          });
        });
      }
    });
  }, 1000);
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
            db.all("SELECT * FROM user_uploads ORDER BY id DESC LIMIT 50", (err, uploads) => {
              db.all("SELECT key, value FROM settings WHERE key LIKE 'template_%'", (err, templateRows) => {
                const templates = {};
                if (templateRows) {
                  templateRows.forEach(r => {
                    templates[r.key.replace('template_', '')] = r.value;
                  });
                }
                db.all("SELECT * FROM materials ORDER BY department, id ASC", (err, allMaterials) => {
                  res.render('admin_super', { 
                    apiKey, 
                    versions: versions || [], 
                    departmentName: '总管理台', 
                    users: users || [], 
                    logs: logs || [], 
                    templates, 
                    uploads: uploads || [], 
                    allMaterials: allMaterials || [] 
                  });
                });
              });
            });
          });
        });
      });
    });
    return;
  }

  // Normal department users
  db.all('SELECT * FROM materials WHERE department = ? ORDER BY id ASC', [req.session.department], (err, materials) => {
    db.all('SELECT * FROM activity_logs WHERE username = ? ORDER BY id DESC LIMIT 20', [req.session.username], (err, logs) => {
      db.all('SELECT * FROM user_uploads WHERE department = ? AND username = ? ORDER BY id DESC LIMIT 20', [req.session.department, req.session.username], (err, uploads) => {
        res.render('admin', { 
          materials, 
          department: req.session.department, 
          departmentName: req.session.departmentName, 
          username: req.session.username,
          logs: logs || [], 
          uploads: uploads || [] 
        });
      });
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

// Init outline with level detection and AI content migration (async)
app.post('/admin/init-outline', isAuthenticated, async (req, res) => {
  if (req.session.department !== 'admin') return res.redirect('/admin');
  
  const { outline, targetDept } = req.body;
  if (!outline || !targetDept) return res.redirect('/admin');

  const lines = outline.split('\n').map(l => l.trim()).filter(l => l.length > 0);

  // Step 1: 获取旧数据并备份
  db.all('SELECT * FROM materials WHERE department = ? ORDER BY id ASC', [targetDept], (err, oldRows) => {
    if (err) return res.redirect('/admin');

    // 备份到版本历史
    if (oldRows && oldRows.length > 0) {
      const snapshot = JSON.stringify(oldRows);
      const tsName = new Date().toLocaleString();
      db.run('INSERT INTO document_versions (department, timestamp, snapshot) VALUES (?, ?, ?)', [targetDept, tsName, snapshot]);
    }

    const hasContent = oldRows && oldRows.some(r => r.content && r.content.length > 0);

    // Step 2: 保存模板
    db.run("INSERT OR REPLACE INTO settings (key, value) VALUES (?, ?)", ['template_' + targetDept, outline], (err) => {
      if (err) return res.redirect('/admin');

      // Step 3: 清空旧数据
      db.run('DELETE FROM materials WHERE department = ?', [targetDept], (err) => {
        if (err) return res.redirect('/admin');

        // Step 4: 批量插入新大纲（使用事务一次性完成）
        let currentParentId = null;
        const insertSQL = [];
        const params = [];
        
        for (const title of lines) {
          let level = 2;
          if (/^[一二三四五六七八九十]+、/.test(title)) level = 1;
          let pId = (level === 1) ? null : currentParentId;
          insertSQL.push('(?, "", ?, ?, ?)');
          params.push(title, targetDept, level, pId);
          if (level === 1) currentParentId = null; // 暂时标记，后续通过查询修正
        }

        db.run('INSERT INTO materials (title, content, department, level, parent_id) VALUES ' + insertSQL.join(','), params, function(err) {
          if (err) return res.redirect('/admin');

          // 修正 parent_id：查询刚插入的数据并建立父子关系
          db.all('SELECT id, title, level FROM materials WHERE department = ? ORDER BY id ASC', [targetDept], (err, newRows) => {
            if (err) return res.redirect('/admin');

            const idMap = [];
            let lastL1 = null;
            newRows.forEach(r => {
              if (r.level === 1) { lastL1 = r.id; }
              idMap.push({ id: r.id, level: r.level, parent: lastL1 });
            });

            // 更新 parent_id
            const updates = idMap.filter(x => x.level === 2 && x.parent).map(x => `WHEN id=${x.id} THEN ${x.parent}`);
            if (updates.length > 0) {
              db.run(`UPDATE materials SET parent_id = CASE ${updates.join(' ')} END WHERE department = ? AND level = 2`, [targetDept]);
            }

            // Step 5: 如果有内容，后台异步AI迁移
            if (hasContent && oldRows) {
              const taskId = 'outline-' + Date.now().toString() + '-' + Math.random().toString(36).substring(2, 7);
              aiTasks[taskId] = { status: 'processing', progress: '正在调用 AI 迁移内容...' };

              migrateContentWithAIAsync(targetDept, oldRows, newRows, taskId, req.session.username);
            }

            db.run("INSERT INTO activity_logs (username, department, action, details) VALUES (?, ?, ?, ?)",
              [req.session.username, targetDept, '重置大纲', `重置大纲共 ${lines.length} 章节${hasContent ? '，原内容正在后台AI迁移中' : '（无内容需迁移）'}，旧版本已归档`]);
            res.redirect('/admin');
          });
        });
      });
    });
  });
});

// AI 内容迁移函数（异步）
async function migrateContentWithAIAsync(targetDept, oldRows, newRows, taskId, username) {
  try {
    const apiKeyRow = await new Promise((resolve, reject) => {
      db.get('SELECT value FROM settings WHERE key = ?', ['deepseek_api_key'], (err, row) => {
        if (err) reject(err); else resolve(row);
      });
    });

    if (!apiKeyRow || !apiKeyRow.value) {
      aiTasks[taskId] = { status: 'completed', message: '无 API Key，跳过AI迁移' };
      return;
    }

    const apiKey = apiKeyRow.value;
    
    // 精简上下文
    const oldContent = oldRows.filter(r => r.content).map(r => {
      const summary = r.content.substring(0, 150).replace(/<[^>]*>/g, '');
      return `【旧: ${r.title}】${summary}${r.content.length > 150 ? '...' : ''}`;
    }).join('\n');

    const newOutline = newRows.map(r => `ID:${r.id} L${r.level} ${r.title}`).join('\n');

    const systemPrompt = `文档内容重组助手。将旧内容分配到新大纲。\n\n旧内容：\n${oldContent}\n\n新大纲：\n${newOutline}\n\n规则：
1. 旧内容匹配新小节：action:"update", id:新ID, content:润色后内容
2. 无法匹配：跳过
3. 保留HTML和图片<img>/视频<video>标签
4. 只返回JSON：[{"action":"update","id":1,"content":"<p>内容</p>"}]`;

    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), 90000);

    const response = await fetch('https://api.deepseek.com/v1/chat/completions', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'Authorization': `Bearer ${apiKey}` },
      signal: controller.signal,
      body: JSON.stringify({
        model: 'deepseek-chat',
        messages: [
          { role: 'system', content: systemPrompt },
          { role: 'user', content: '请开始重组内容' }
        ],
        temperature: 0.7
      })
    });

    clearTimeout(timeoutId);

    if (!response.ok) {
      aiTasks[taskId] = { status: 'failed', error: 'AI 请求失败' };
      return;
    }

    const data = await response.json();
    let text = data.choices[0].message.content.trim();
    if (text.startsWith('```json')) text = text.substring(7, text.length - 3).trim();
    else if (text.startsWith('```')) text = text.substring(3, text.length - 3).trim();

    const updates = JSON.parse(text);

    await new Promise((resolve, reject) => {
      const updateStmt = db.prepare('UPDATE materials SET content = ? WHERE id = ? AND department = ?');
      db.serialize(() => {
        db.run('BEGIN TRANSACTION');
        for (const up of updates) {
          if (up.action === 'update' && up.id && up.content) updateStmt.run(up.content, up.id, targetDept);
        }
        db.run('COMMIT', () => {
          updateStmt.finalize();
          resolve();
        });
      });
    });

    aiTasks[taskId] = { status: 'completed', message: 'AI 内容迁移完成' };
  } catch (err) {
    console.error('AI 迁移失败:', err.message);
    aiTasks[taskId] = { status: 'failed', error: err.message.substring(0, 200) };
  }
}

// AI Generate content & Save to MD (async)
app.post('/admin/generate-ai-content', isAuthenticated, async (req, res) => {
  const { newMaterials } = req.body;
  const dept = req.session.department;

  if (newMaterials && newMaterials.length > 50000) {
    return res.status(400).json({ error: '上传内容过大（限5万字以内），请分次处理以确保生成质量。' });
  }

  const taskId = Date.now().toString() + '-' + Math.random().toString(36).substring(2, 7);
  aiTasks[taskId] = { status: 'processing', progress: '正在调用 AI...' };

  db.get('SELECT value FROM settings WHERE key = ?', ['deepseek_api_key'], async (err, row) => {
    const apiKey = row ? row.value : null;
    if (!apiKey) {
      aiTasks[taskId] = { status: 'failed', error: '请先在系统设置中配置 API Key' };
      return;
    }

    // 只获取标题和层级，减少上下文
    db.all('SELECT id, title, level, parent_id, content FROM materials WHERE department = ? ORDER BY id ASC', [dept], async (err, rows) => {
      if (err || rows.length === 0) {
        aiTasks[taskId] = { status: 'failed', error: '当前部门大纲为空，请先初始化大纲。' };
        return;
      }

      // 精简上下文：只发送标题+内容摘要
      const outlineContext = rows.map(r => {
        const contentLen = r.content ? r.content.length : 0;
        const summary = r.content ? r.content.substring(0, 100).replace(/<[^>]*>/g, '') : '空';
        return `ID:${r.id} L${r.level} ${r.title} [内容长度:${contentLen}字, 摘要:${summary}${contentLen > 100 ? '...' : ''}]`;
      }).join('\n');

      let systemPrompt = `智能文档助手，将新资料归类到文档结构。\n大纲：\n${outlineContext}\n\n规则：
1. 知识点已存在匹配小节：action:"overwrite" 覆盖该小节，融合去重优化
2. 与某大章节有关联但无匹配小节：action:"create" 指定 parent_id(所属大章节ID) 和 title(你自动生成的二级标题)，并写入内容
3. 完全独立新事项但属于文档范畴：action:"append" 追加到最相近小节
4. 与本文档完全无关：action:"reject" 说明原因
5. 必须保留原有的<img>和<video>标签！

返回JSON数组：
[{"action":"overwrite","id":1,"content":"<p>内容</p>"}]
[{"action":"create","parent_id":3,"title":"新二级标题","content":"<p>内容</p>"}]
[{"action":"reject","reason":"与本文档无关的原因"}]
只输出JSON！`;

      const controller = new AbortController();
      const timeoutId = setTimeout(() => controller.abort(), 90000);

      try {
        aiTasks[taskId].progress = 'AI 正在处理中...';
        const response = await fetch('https://api.deepseek.com/v1/chat/completions', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json', 'Authorization': `Bearer ${apiKey}` },
          signal: controller.signal,
          body: JSON.stringify({
            model: 'deepseek-v4-pro',
            messages: [
              { role: 'system', content: systemPrompt },
              { role: 'user', content: `富文本新资料：\n${newMaterials}` }
            ],
            temperature: 0.7
          })
        });

        clearTimeout(timeoutId);

        if (!response.ok) throw new Error("AI API Error: " + await response.text());

        const data = await response.json();
        aiTasks[taskId].progress = 'AI 处理完成，正在保存...';

        let text = data.choices[0].message.content.trim();
        if (text.startsWith('```json')) text = text.substring(7, text.length - 3).trim();
        else if (text.startsWith('```')) text = text.substring(3, text.length - 3).trim();

        const updates = JSON.parse(text);

        const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
        const mdFileName = `doc_${dept}_${timestamp}.md`;
        let mdContent = `# 新增内容归档 - ${new Date().toLocaleString()}\n\n## 原始输入:\n${newMaterials}\n\n---\n\n## 归类结果:\n\n`;

        const rejectedItems = updates.filter(u => u.action === 'reject');
        const validUpdates = updates.filter(u => u.action !== 'reject');

        for (const up of updates) {
          if (up.action === 'reject') {
            mdContent += `### [无关] 已跳过：${up.reason}\n\n`;
          } else {
            const t = rows.find(r => r.id == up.id)?.title || rows.find(r => r.id == up.parent_id)?.title || '未知';
            mdContent += `### [${up.action || 'update'}] ${t} (ID:${up.id || up.parent_id})\n${up.content || up.title}\n\n`;
          }
        }
        fs.writeFileSync(path.join(__dirname, 'public/docs', mdFileName), mdContent);

        db.run('INSERT INTO user_uploads (filename, original_name, file_path, file_type, department, username) VALUES (?, ?, ?, ?, ?, ?)',
          [mdFileName, `AI提交归档_${timestamp}.md`, '/docs/' + mdFileName, 'ai_archive', dept, req.session.username]);

        const updateStmt = db.prepare('UPDATE materials SET content = CASE WHEN content IS NULL OR content = "" THEN ? ELSE content || "<br><br>" || ? END WHERE id = ? AND department = ?');
        const overwriteStmt = db.prepare('UPDATE materials SET content = ? WHERE id = ? AND department = ?');
        const createStmt = db.prepare('INSERT INTO materials (title, content, department, level, parent_id) VALUES (?, ?, ?, 2, ?)');

        db.serialize(() => {
          db.run('BEGIN TRANSACTION');
          for (const up of validUpdates) {
            if (up.action === 'append' && up.id && up.content) updateStmt.run(up.content, up.content, up.id, dept);
            else if (up.action === 'overwrite' && up.id && up.content) overwriteStmt.run(up.content, up.id, dept);
            else if (up.action === 'create' && up.parent_id && up.title && up.content) createStmt.run(up.title, up.content, dept, up.parent_id);
            else if (up.id && up.content && !up.action) updateStmt.run(up.content, up.content, up.id, dept);
          }
          db.run('COMMIT', () => {
            updateStmt.finalize(); overwriteStmt.finalize(); createStmt.finalize();

            db.all('SELECT * FROM materials WHERE department = ? ORDER BY id ASC', [dept], (err, currentRows) => {
              if (!err && currentRows) {
                db.run('INSERT INTO document_versions (department, timestamp, snapshot) VALUES (?, ?, ?)',
                  [dept, new Date().toLocaleString(), JSON.stringify(currentRows)]);
              }
            });

            const rejectMsg = rejectedItems.length > 0 ? `\n⚠️ 以下 ${rejectedItems.length} 项被跳过：${rejectedItems.map(r => r.reason).join('；')}` : '';

            db.run("INSERT INTO activity_logs (username, department, action, details) VALUES (?, ?, ?, ?)",
              [req.session.username, dept, 'AI智能填充', `成功处理新资料，归档: ${mdFileName}${rejectMsg}`]);
            
            aiTasks[taskId] = { 
              status: 'completed', 
              message: `AI 处理并归档成功${rejectedItems.length > 0 ? '（部分内容与文档无关，已跳过）' : ''}`, 
              mdFile: mdFileName,
              rejected: rejectedItems.length > 0 ? rejectedItems.map(r => r.reason) : []
            };
          });
        });
      } catch (err) {
        clearTimeout(timeoutId);
        console.error("AI Generation Error:", err);
        const errorDetail = err.name === 'AbortError' ? 'AI 请求超时（90秒），请尝试分段处理。' : err.message;
        db.run("INSERT INTO activity_logs (username, department, action, details) VALUES (?, ?, ?, ?)",
          [req.session.username, dept, 'AI处理失败', `错误: ${errorDetail.substring(0, 200)}`]);
        aiTasks[taskId] = { status: 'failed', error: errorDetail };
      }
    });
  });

  // 立即返回 taskId
  res.json({ success: true, taskId: taskId, message: 'AI 处理已启动，请稍候...' });
});

// 查询 AI 任务状态
app.get('/admin/ai-task-status/:taskId', isAuthenticated, (req, res) => {
  const task = aiTasks[req.params.taskId];
  if (!task) return res.json({ status: 'not_found' });
  res.json(task);
  // 完成后清理
  if (task.status === 'completed' || task.status === 'failed') {
    setTimeout(() => delete aiTasks[req.params.taskId], 60000);
  }
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
  const media_path = '/uploads/' + req.file.filename;
  const media_type = req.file.mimetype.startsWith('image/') ? 'image' : (req.file.mimetype.startsWith('video/') ? 'video' : null);
  
  // Log upload
  db.run('INSERT INTO user_uploads (filename, original_name, file_path, file_type, department, username) VALUES (?, ?, ?, ?, ?, ?)',
    [req.file.filename, req.file.originalname, media_path, media_type, req.session.department, req.session.username]);

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

  // Log parsed upload
  db.run('INSERT INTO user_uploads (filename, original_name, file_path, file_type, department, username) VALUES (?, ?, ?, ?, ?, ?)',
    [req.file.filename, req.file.originalname, '/uploads/' + req.file.filename, 'document', req.session.department, req.session.username]);

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

app.listen(port, '0.0.0.0', () => {
  console.log(`Document server is running!`);
  console.log(`- Local Access: http://localhost:${port}`);
  
  const os = require('os');
  const nets = os.networkInterfaces();
  for (const name of Object.keys(nets)) {
    for (const net of nets[name]) {
      if (net.family === 'IPv4' && !net.internal) {
        console.log(`- Network Access: http://${net.address}:${port}`);
      }
    }
  }
});
