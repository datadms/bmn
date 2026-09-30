/**
 * GAS-Bloger V2 II Backend Engine
 * Menangani routing GET dan POST dengan keamanan CORS dan JSON Response murni.
 */

function doGet(e) {
  return handleRequest(e, 'GET');
}

function doPost(e) {
  return handleRequest(e, 'POST');
}

function handleRequest(e, method) {
  try {
    const params = method === 'GET' ? e.parameter : JSON.parse(e.postData.contents || "{}");
    const action = params.action || 'getMenus';
    
    let result = { status: 'error', message: 'Aksi tidak dikenal.' };

    if (action === 'getMenus') {
      result = getMenusData(params.userId, params.role);
    } else if (action === 'login') {
      result = authenticateUser(params.username, params.password);
    } else if (action === 'saveMenu') {
      result = saveMenuData(params);
    } else if (action === 'deleteMenu') {
      result = deleteMenuData(params.id);
    } else if (action === 'getUsers') {
      result = getAllUsers();
    } else if (action === 'saveUser') {
      result = saveUserData(params);
    } else if (action === 'deleteUser') {
      result = deleteUserData(params.userId);
    }

    return createJsonResponse(result);
  } catch (err) {
    return createJsonResponse({
      status: 'error',
      message: 'Kesalahan Server Internal: ' + err.toString()
    });
  }
}

function createJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (name === 'Menus') {
      sheet.appendRow(['id', 'parentId', 'title', 'category', 'visibility', 'type', 'iconClass', 'targetUrl', 'description', 'ownerUserId']);
    } else if (name === 'Users') {
      sheet.appendRow(['userId', 'username', 'password', 'fullName', 'role']);
      // Default Admin
      sheet.appendRow(['ADM-01', 'admin', 'admin123', 'Super Administrator', 'ADMIN']);
    }
  }
  return sheet;
}

function getMenusData(userId, role) {
  const sheet = getSheet('Menus');
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return { status: 'success', data: [] };

  const headers = rows[0];
  const dataRows = rows.slice(1);
  
  let allMenus = dataRows.map(row => {
    let obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });

  // Filter visibility
  let filtered = allMenus.filter(item => {
    if (item.visibility === 'UMUM') return true;
    if (role === 'ADMIN') return true;
    if (item.ownerUserId && item.ownerUserId === userId) return true;
    return false;
  });

  // Build tree structure (Root & Submenus)
  let rootMenus = filtered.filter(item => !item.parentId || item.parentId === 'ROOT');
  rootMenus.forEach(root => {
    root.submenus = filtered.filter(sub => sub.parentId === root.id);
  });

  return { status: 'success', data: rootMenus };
}

function authenticateUser(username, password) {
  const sheet = getSheet('Users');
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][1] === username && rows[i][2] === password) {
      return {
        status: 'success',
        user: {
          userId: rows[i][0],
          username: rows[i][1],
          fullName: rows[i][3],
          role: rows[i][4]
        }
      };
    }
  }
  return { status: 'error', message: 'Username atau password salah!' };
}

function saveMenuData(p) {
  const sheet = getSheet('Menus');
  const rows = sheet.getDataRange().getValues();
  let id = p.id;
  
  if (!id) {
    id = 'MN-' + Date.now();
    sheet.appendRow([id, p.parentId || 'ROOT', p.title, p.category, p.visibility, p.type, p.iconClass, p.targetUrl, p.description, p.ownerUserId]);
  } else {
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === id) {
        sheet.getRange(i + 1, 2, 1, 9).setValues([[
          p.parentId || 'ROOT', p.title, p.category, p.visibility, p.type, p.iconClass, p.targetUrl, p.description, p.ownerUserId
        ]]);
        break;
      }
    }
  }
  return { status: 'success', message: 'Data menu berhasil disimpan.' };
}

function deleteMenuData(id) {
  const sheet = getSheet('Menus');
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === id) {
      sheet.deleteRow(i + 1);
      return { status: 'success', message: 'Menu berhasil dihapus.' };
    }
  }
  return { status: 'error', message: 'Menu tidak ditemukan.' };
}

function getAllUsers() {
  const sheet = getSheet('Users');
  const rows = sheet.getDataRange().getValues();
  const users = rows.slice(1).map(r => ({
    userId: r[0],
    username: r[1],
    fullName: r[3],
    role: r[4]
  }));
  return { status: 'success', data: users };
}

function saveUserData(p) {
  const sheet = getSheet('Users');
  const rows = sheet.getDataRange().getValues();
  let found = false;
  
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === p.userId) {
      sheet.getRange(i + 1, 2, 1, 4).setValues([[p.username, p.password, p.fullName, p.role]]);
      found = true;
      break;
    }
  }
  if (!found) {
    sheet.appendRow([p.userId, p.username, p.password, p.fullName, p.role]);
  }
  return { status: 'success', message: 'Data user berhasil disimpan.' };
}

function deleteUserData(userId) {
  const sheet = getSheet('Users');
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === userId) {
      sheet.deleteRow(i + 1);
      return { status: 'success', message: 'User berhasil dihapus.' };
    }
  }
  return { status: 'error', message: 'User tidak ditemukan.' };
}
