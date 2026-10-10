/**
 * CLOUDDOCS - ONLINE REALTIME SYNC & FILE EXPLORER ENGINE
 * Multi-device bidirectional synchronization via Firebase Realtime Database,
 * Offline-first localStorage caching, Folder & Docx management, and QR device pairing.
 */

class CloudDocsEngine {
  constructor() {
    this.storageKey = 'clouddocs_data';
    this.userKey = 'clouddocs_active_user';
    this.firebaseDbUrl = 'https://vocalearn-app-default-rtdb.asia-southeast1.firebasedatabase.app';
    
    // User Identity
    this.activeUser = this.detectUser();
    this.currentFolderId = null; // null = root
    this.activeDocId = null;
    this.viewMode = localStorage.getItem('clouddocs_view_mode') || 'grid'; // 'grid' | 'list'
    this.sortBy = 'newest'; // 'newest' | 'name' | 'size'
    this.searchQuery = '';
    
    // State
    this.folders = [];
    this.documents = [];
    this.lastSyncTimestamp = 0;
    this.isOnline = navigator.onLine;
    this.pushTimer = null;
    this.pollInterval = null;
    this.firebaseDb = null;
    
    // Initial data load
    this.loadLocalData();
    this.initFirebase();
    this.initEventBindings();
    this.startPeriodicSync();
  }

  // Detect active user from URL query or localStorage or generate default
  detectUser() {
    const urlParams = new URLSearchParams(window.location.search);
    const userFromUrl = urlParams.get('user') || urlParams.get('acc');
    if (userFromUrl) {
      const cleanUser = userFromUrl.trim().toLowerCase().replace(/[^a-z0-9_-]/g, '');
      if (cleanUser) {
        localStorage.setItem(this.userKey, cleanUser);
        return cleanUser;
      }
    }
    const savedUser = localStorage.getItem(this.userKey);
    if (savedUser) return savedUser;
    
    // Generate friendly default username
    const defaultUser = 'user_' + Math.random().toString(36).substring(2, 7);
    localStorage.setItem(this.userKey, defaultUser);
    return defaultUser;
  }

  // =========================================================================
  // 1. DATA PERSISTENCE & SAMPLE INITIALIZATION
  // =========================================================================
  loadLocalData() {
    try {
      const raw = localStorage.getItem(`${this.storageKey}_${this.activeUser}`);
      if (raw) {
        const parsed = JSON.parse(raw);
        this.folders = Array.isArray(parsed.folders) ? parsed.folders : [];
        this.documents = Array.isArray(parsed.documents) ? parsed.documents : [];
        this.lastSyncTimestamp = parsed.updatedAt || 0;
        return;
      }
    } catch (e) {
      console.warn('Error reading local data:', e);
    }

    // Default Starter Content if empty
    this.folders = [
      {
        id: 'folder_work',
        name: 'Công việc & Dự án',
        parentId: null,
        color: '#185abd',
        createdAt: Date.now() - 86400000,
        updatedAt: Date.now() - 86400000
      },
      {
        id: 'folder_personal',
        name: 'Tài liệu cá nhân',
        parentId: null,
        color: '#107c41',
        createdAt: Date.now() - 43200000,
        updatedAt: Date.now() - 43200000
      }
    ];

    this.documents = [
      {
        id: 'doc_welcome',
        title: 'Hướng dẫn sử dụng & Đồng bộ đa thiết bị.docx',
        folderId: 'folder_work',
        content: `
          <h1 style="color: #185abd; font-size: 24pt; margin-bottom: 12pt;">HƯỚNG DẪN ĐỒNG BỘ ĐA THIẾT BỊ ONLINE</h1>
          <p style="font-size: 11pt; line-height: 1.6; margin-bottom: 10pt;">
            Chào mừng bạn đến với <strong>CloudDocs</strong> - Trình soạn thảo văn bản và Quản lý tệp Word chuyên nghiệp trực tuyến!
          </p>
          <div style="background-color: #f0f4f9; border-left: 4px solid #185abd; padding: 12px 16px; margin: 16px 0; border-radius: 4px;">
            <strong>Mẹo đồng bộ nhanh:</strong> Bạn có thể mở tài liệu này trên <em>Điện thoại di động</em> bằng cách bấm nút <strong>"Mở trên ĐT"</strong> ở góc trên bên phải để quét mã QR hoặc sao chép liên kết. Mọi nội dung gõ trên điện thoại hay laptop đều tự động lưu và cập nhật trực tiếp theo thời gian thực!
          </div>
          <h2 style="color: #103f82; font-size: 16pt; margin-top: 16pt; margin-bottom: 8pt;">Các tính năng nổi bật:</h2>
          <ul style="font-size: 11pt; line-height: 1.8;">
            <li><strong>Quản lý thư mục:</strong> Tạo thư mục con, đổi màu sắc, phân loại tài liệu khoa học.</li>
            <li><strong>Soạn thảo Word phong phú:</strong> Đầy đủ thanh công cụ Ribbon, Font chữ, In đậm, Căn lề, Chèn bảng, Chèn ảnh, Thước đo A4.</li>
            <li><strong>Đồng bộ tức thì (Realtime):</strong> Gõ trên điện thoại, máy tính thấy ngay tức thì không cần tải lại trang.</li>
            <li><strong>Tải xuống & Tải lên (.docx):</strong> Hỗ trợ xuất và nhập file Microsoft Word thực thụ.</li>
          </ul>
        `,
        preview: 'Chào mừng bạn đến với CloudDocs - Trình soạn thảo văn bản và Quản lý tệp Word chuyên nghiệp trực tuyến...',
        wordCount: 145,
        pageCount: 1,
        createdAt: Date.now() - 3600000,
        updatedAt: Date.now() - 3600000
      },
      {
        id: 'doc_plan',
        title: 'Kế hoạch công tác năm 2026.docx',
        folderId: null, // Root
        content: `
          <h1 style="color: #185abd; text-align: center; font-size: 22pt;">KẾ HOẠCH CÔNG TÁC VÀ MỤC TIÊU</h1>
          <p style="text-align: center; font-style: italic; color: #666;">Được lập ngày: 10/10/2026</p>
          <hr style="margin: 16px 0; border: none; border-top: 1px solid #ccc;">
          <h3 style="color: #103f82;">1. Mục tiêu trọng tâm</h3>
          <p>Tối ưu hóa quy trình làm việc từ xa, đồng bộ dữ liệu liên tục trên laptop và điện thoại.</p>
          <h3 style="color: #103f82;">2. Tiến độ triển khai</h3>
          <p>Hoàn thành giai đoạn kiểm thử và bàn giao sản phẩm trước quý 4.</p>
        `,
        preview: 'Kế hoạch công tác và mục tiêu trọng tâm: Tối ưu hóa quy trình làm việc từ xa...',
        wordCount: 68,
        pageCount: 1,
        createdAt: Date.now() - 7200000,
        updatedAt: Date.now() - 7200000
      }
    ];

    this.saveLocalData();
  }

  saveLocalData() {
    try {
      const payload = {
        folders: this.folders,
        documents: this.documents,
        updatedAt: Date.now(),
        syncUser: this.activeUser
      };
      localStorage.setItem(`${this.storageKey}_${this.activeUser}`, JSON.stringify(payload));
    } catch (e) {
      console.warn('Error saving to localStorage:', e);
    }
  }

  // =========================================================================
  // 2. FIREBASE REALTIME DATABASE SYNC
  // =========================================================================
  initFirebase() {
    try {
      if (window.firebase && !window.firebase.apps.length) {
        window.firebase.initializeApp({
          databaseURL: this.firebaseDbUrl
        });
      }
      if (window.firebase && window.firebase.database) {
        this.firebaseDb = window.firebase.database();
        this.listenRealtimeFirebase();
      }
    } catch (e) {
      console.log('Firebase SDK init skipped, using high-speed REST sync:', e);
    }

    // Always fetch once on start
    this.syncFromCloud();
  }

  listenRealtimeFirebase() {
    if (!this.firebaseDb || !this.activeUser) return;
    try {
      const userRef = this.firebaseDb.ref(`clouddocs_users/${this.activeUser}`);
      userRef.on('value', (snapshot) => {
        const val = snapshot.val();
        if (val && val.updatedAt && val.updatedAt > this.lastSyncTimestamp) {
          console.log('[Realtime Sync] Nhận dữ liệu mới từ Cloud!');
          this.applyCloudData(val);
        }
      });
      this.updateSyncBadge('online');
    } catch (err) {
      console.warn('Realtime listener fallback to REST:', err);
    }
  }

  async syncFromCloud() {
    if (!this.activeUser) return;
    const endpoint = `${this.firebaseDbUrl}/clouddocs_users/${encodeURIComponent(this.activeUser)}.json`;
    try {
      const res = await fetch(endpoint, { method: 'GET' });
      if (res.ok) {
        const cloudData = await res.json();
        if (cloudData && cloudData.updatedAt) {
          if (cloudData.updatedAt > this.lastSyncTimestamp) {
            this.applyCloudData(cloudData);
          } else if (this.lastSyncTimestamp > cloudData.updatedAt) {
            // Local is newer, push to cloud
            this.pushToCloud();
          }
        } else if (!cloudData) {
          // Cloud empty for this user, upload starter data
          this.pushToCloud();
        }
        this.updateSyncBadge('online');
      }
    } catch (e) {
      this.updateSyncBadge(navigator.onLine ? 'online' : 'offline');
    }
  }

  applyCloudData(cloudData) {
    this.folders = Array.isArray(cloudData.folders) ? cloudData.folders : [];
    this.documents = Array.isArray(cloudData.documents) ? cloudData.documents : [];
    this.lastSyncTimestamp = cloudData.updatedAt || Date.now();
    this.saveLocalData();
    
    // Refresh Explorer View
    this.renderExplorer();
    
    // If active document is open in editor, check if content was updated remotely
    if (this.activeDocId) {
      const currentDoc = this.documents.find(d => d.id === this.activeDocId);
      if (currentDoc) {
        const editorContentEl = document.getElementById('document-content');
        const editorTitleEl = document.getElementById('doc-title-input');
        if (editorContentEl && document.activeElement !== editorContentEl) {
          // Don't overwrite if user is currently typing
          editorContentEl.innerHTML = currentDoc.content || '';
          if (editorTitleEl) editorTitleEl.value = currentDoc.title;
          this.showToast('Đã nhận nội dung cập nhật mới từ thiết bị khác!', 'fa-solid fa-arrows-rotate');
        }
      }
    }
  }

  triggerDataChange(immediate = false) {
    this.saveLocalData();
    this.renderExplorer();
    
    clearTimeout(this.pushTimer);
    if (immediate) {
      this.pushToCloud();
    } else {
      this.updateSyncBadge('syncing');
      this.pushTimer = setTimeout(() => this.pushToCloud(), 600);
    }
  }

  async pushToCloud() {
    if (!this.activeUser) return;
    const now = Date.now();
    this.lastSyncTimestamp = now;

    const payload = {
      folders: this.folders,
      documents: this.documents,
      updatedAt: now,
      syncUser: this.activeUser
    };

    this.updateSyncBadge('syncing');

    // 1. Try Firebase SDK if available
    if (this.firebaseDb) {
      try {
        await this.firebaseDb.ref(`clouddocs_users/${this.activeUser}`).set(payload);
        this.updateSyncBadge('online');
        return;
      } catch (err) {}
    }

    // 2. High-speed REST PUT fallback
    try {
      const endpoint = `${this.firebaseDbUrl}/clouddocs_users/${encodeURIComponent(this.activeUser)}.json`;
      const res = await fetch(endpoint, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });
      if (res.ok) {
        this.updateSyncBadge('online');
      } else {
        this.updateSyncBadge('offline');
      }
    } catch (e) {
      this.updateSyncBadge('offline');
    }
  }

  startPeriodicSync() {
    // Poll every 3 seconds to ensure mobile <-> laptop instantaneous sync
    clearInterval(this.pollInterval);
    this.pollInterval = setInterval(() => {
      if (navigator.onLine) {
        this.syncFromCloud();
      }
    }, 3500);

    window.addEventListener('online', () => {
      this.updateSyncBadge('online');
      this.pushToCloud();
      this.showToast('Đã kết nối Internet! Đang đồng bộ...');
    });

    window.addEventListener('offline', () => {
      this.updateSyncBadge('offline');
      this.showToast('Bạn đang ở chế độ ngoại tuyến. Dữ liệu vẫn được lưu vào máy.');
    });
  }

  updateSyncBadge(status) {
    const badge = document.getElementById('explorer-sync-badge');
    const label = document.getElementById('explorer-sync-text');
    const editorSaveStatus = document.getElementById('save-status');

    if (badge) {
      badge.className = `cloud-sync-badge ${status}`;
    }

    if (label) {
      if (status === 'online') label.textContent = 'Đồng bộ Trực tuyến (Cloud)';
      else if (status === 'syncing') label.textContent = 'Đang đồng bộ...';
      else label.textContent = 'Lưu cục bộ (Offline)';
    }

    if (editorSaveStatus) {
      if (status === 'online') {
        editorSaveStatus.innerHTML = '<i class="fa-solid fa-cloud-check text-green"></i> Đã đồng bộ Cloud';
      } else if (status === 'syncing') {
        editorSaveStatus.innerHTML = '<i class="fa-solid fa-arrows-rotate fa-spin text-blue"></i> Đang lưu lên Cloud...';
      } else {
        editorSaveStatus.innerHTML = '<i class="fa-solid fa-hard-drive"></i> Đã lưu vào máy (Offline)';
      }
    }
  }

  // =========================================================================
  // 3. FOLDER & DOCUMENT CRUD OPERATIONS
  // =========================================================================
  createFolder(name, color = '#185abd') {
    const trimmed = name.trim();
    if (!trimmed) return;
    const newFolder = {
      id: 'f_' + Date.now() + '_' + Math.random().toString(36).substring(2, 6),
      name: trimmed,
      parentId: this.currentFolderId,
      color: color,
      createdAt: Date.now(),
      updatedAt: Date.now()
    };
    this.folders.push(newFolder);
    this.triggerDataChange(true);
    this.showToast(`Đã tạo thư mục "${trimmed}"`);
  }

  renameFolder(folderId, newName) {
    const trimmed = newName.trim();
    if (!trimmed) return;
    const folder = this.folders.find(f => f.id === folderId);
    if (folder) {
      folder.name = trimmed;
      folder.updatedAt = Date.now();
      this.triggerDataChange(true);
      this.showToast(`Đã đổi tên thư mục thành "${trimmed}"`);
    }
  }

  deleteFolder(folderId) {
    const folder = this.folders.find(f => f.id === folderId);
    if (!folder) return;

    if (!confirm(`Bạn có chắc muốn xóa thư mục "${folder.name}" và các tệp bên trong?`)) {
      return;
    }

    // Delete folder and nested folders/documents recursively
    const folderIdsToDelete = new Set([folderId]);
    let added = true;
    while (added) {
      added = false;
      this.folders.forEach(f => {
        if (folderIdsToDelete.has(f.parentId) && !folderIdsToDelete.has(f.id)) {
          folderIdsToDelete.add(f.id);
          added = true;
        }
      });
    }

    this.folders = this.folders.filter(f => !folderIdsToDelete.has(f.id));
    this.documents = this.documents.filter(d => !folderIdsToDelete.has(d.folderId));

    if (folderIdsToDelete.has(this.currentFolderId)) {
      this.currentFolderId = null;
    }

    this.triggerDataChange(true);
    this.showToast(`Đã xóa thư mục "${folder.name}"`);
  }

  createDocument(title = 'Tài liệu mới.docx', initialContent = '') {
    let cleanTitle = title.trim();
    if (!cleanTitle.toLowerCase().endsWith('.docx')) {
      cleanTitle += '.docx';
    }

    const defaultContent = initialContent || `
      <h1 style="color: #185abd; font-size: 22pt;">${cleanTitle.replace('.docx', '')}</h1>
      <p style="font-size: 11pt; line-height: 1.6;">Bắt đầu soạn thảo văn bản tại đây...</p>
    `;

    const newDoc = {
      id: 'doc_' + Date.now() + '_' + Math.random().toString(36).substring(2, 6),
      title: cleanTitle,
      folderId: this.currentFolderId,
      content: defaultContent,
      preview: this.extractPreviewText(defaultContent),
      wordCount: 10,
      pageCount: 1,
      createdAt: Date.now(),
      updatedAt: Date.now()
    };

    this.documents.unshift(newDoc);
    this.triggerDataChange(true);
    this.openDocument(newDoc.id);
    this.showToast(`Đã tạo "${cleanTitle}"`);
  }

  openDocument(docId) {
    const doc = this.documents.find(d => d.id === docId);
    if (!doc) return;

    this.activeDocId = docId;
    const docContentEl = document.getElementById('document-content');
    const docTitleInput = document.getElementById('doc-title-input');

    if (docContentEl) docContentEl.innerHTML = doc.content || '';
    if (docTitleInput) docTitleInput.value = doc.title || 'Tài liệu 1.docx';

    // Switch View to Word Editor
    this.switchView('editor');

    // Trigger metrics update
    if (window.updateWordMetrics) {
      window.updateWordMetrics();
    }
  }

  saveActiveDocument(forceImmediate = false) {
    if (!this.activeDocId) return;
    const doc = this.documents.find(d => d.id === this.activeDocId);
    if (!doc) return;

    const docContentEl = document.getElementById('document-content');
    const docTitleInput = document.getElementById('doc-title-input');

    if (docContentEl) {
      doc.content = docContentEl.innerHTML;
      doc.preview = this.extractPreviewText(docContentEl.innerHTML);
      const text = docContentEl.innerText || '';
      doc.wordCount = text.trim() ? text.trim().split(/\s+/).length : 0;
    }

    if (docTitleInput && docTitleInput.value.trim()) {
      let t = docTitleInput.value.trim();
      if (!t.toLowerCase().endsWith('.docx') && !t.includes('.')) {
        t += '.docx';
      }
      doc.title = t;
    }

    doc.updatedAt = Date.now();
    this.triggerDataChange(forceImmediate);
  }

  renameDocument(docId, newTitle) {
    let trimmed = newTitle.trim();
    if (!trimmed) return;
    if (!trimmed.toLowerCase().endsWith('.docx') && !trimmed.includes('.')) {
      trimmed += '.docx';
    }
    const doc = this.documents.find(d => d.id === docId);
    if (doc) {
      doc.title = trimmed;
      doc.updatedAt = Date.now();
      this.triggerDataChange(true);
      this.showToast(`Đã đổi tên thành "${trimmed}"`);
    }
  }

  deleteDocument(docId) {
    const doc = this.documents.find(d => d.id === docId);
    if (!doc) return;
    if (!confirm(`Bạn có chắc muốn xóa tệp "${doc.title}"?`)) return;

    this.documents = this.documents.filter(d => d.id !== docId);
    if (this.activeDocId === docId) {
      this.activeDocId = null;
      this.switchView('explorer');
    }
    this.triggerDataChange(true);
    this.showToast(`Đã xóa "${doc.title}"`);
  }

  duplicateDocument(docId) {
    const doc = this.documents.find(d => d.id === docId);
    if (!doc) return;

    const copyDoc = {
      ...doc,
      id: 'doc_' + Date.now() + '_' + Math.random().toString(36).substring(2, 6),
      title: doc.title.replace('.docx', '') + ' (Bản sao).docx',
      createdAt: Date.now(),
      updatedAt: Date.now()
    };

    this.documents.unshift(copyDoc);
    this.triggerDataChange(true);
    this.showToast(`Đã nhân bản "${copyDoc.title}"`);
  }

  moveDocument(docId, targetFolderId) {
    const doc = this.documents.find(d => d.id === docId);
    if (!doc) return;
    doc.folderId = targetFolderId;
    doc.updatedAt = Date.now();
    this.triggerDataChange(true);
    this.showToast(`Đã di chuyển "${doc.title}"`);
  }

  extractPreviewText(html) {
    const tmp = document.createElement('div');
    tmp.innerHTML = html;
    const txt = tmp.innerText || tmp.textContent || '';
    return txt.replace(/\s+/g, ' ').trim().substring(0, 140);
  }

  // =========================================================================
  // 4. RENDERING EXPLORER UI
  // =========================================================================
  renderExplorer() {
    this.renderHeader();
    this.renderBreadcrumbs();
    this.renderFolders();
    this.renderDocuments();
  }

  renderHeader() {
    const userLabel = document.getElementById('user-account-name');
    const userAvatar = document.getElementById('user-account-avatar');
    if (userLabel) userLabel.textContent = this.activeUser;
    if (userAvatar) userAvatar.textContent = (this.activeUser[0] || 'U').toUpperCase();
  }

  renderBreadcrumbs() {
    const container = document.getElementById('breadcrumb-list');
    if (!container) return;

    const path = [];
    let curId = this.currentFolderId;
    while (curId) {
      const f = this.folders.find(x => x.id === curId);
      if (f) {
        path.unshift(f);
        curId = f.parentId;
      } else {
        break;
      }
    }

    let html = `
      <span class="breadcrumb-item ${this.currentFolderId === null ? 'active' : ''}" data-folder-id="root">
        <i class="fa-solid fa-house"></i> Tất cả tệp
      </span>
    `;

    path.forEach((f, idx) => {
      const isLast = idx === path.length - 1;
      html += `
        <i class="fa-solid fa-chevron-right breadcrumb-separator"></i>
        <span class="breadcrumb-item ${isLast ? 'active' : ''}" data-folder-id="${f.id}">
          <i class="fa-solid fa-folder" style="color: ${f.color || '#f59e0b'};"></i> ${this.escapeHtml(f.name)}
        </span>
      `;
    });

    container.innerHTML = html;

    // Attach click events
    container.querySelectorAll('.breadcrumb-item').forEach(el => {
      el.addEventListener('click', () => {
        const fid = el.dataset.folderId;
        this.currentFolderId = fid === 'root' ? null : fid;
        this.renderExplorer();
      });
    });
  }

  renderFolders() {
    const grid = document.getElementById('folders-grid');
    const section = document.getElementById('folders-section');
    const countEl = document.getElementById('folders-count');
    if (!grid || !section) return;

    // Filter by current directory and search query
    let currentFolders = this.folders.filter(f => f.parentId === this.currentFolderId);
    if (this.searchQuery) {
      const q = this.searchQuery.toLowerCase();
      currentFolders = this.folders.filter(f => f.name.toLowerCase().includes(q));
    }

    if (countEl) countEl.textContent = `(${currentFolders.length})`;

    if (currentFolders.length === 0) {
      if (this.searchQuery) {
        section.style.display = 'none';
      } else {
        grid.innerHTML = '';
        section.style.display = 'none';
      }
      return;
    }

    section.style.display = 'block';

    grid.innerHTML = currentFolders.map(f => {
      const docCount = this.documents.filter(d => d.folderId === f.id).length;
      return `
        <div class="folder-card" data-folder-id="${f.id}">
          <div class="folder-card-main">
            <i class="fa-solid fa-folder folder-icon" style="color: ${f.color || '#f59e0b'};"></i>
            <div class="folder-info">
              <span class="folder-name" title="${this.escapeHtml(f.name)}">${this.escapeHtml(f.name)}</span>
              <span class="folder-meta">${docCount} tệp docx</span>
            </div>
          </div>
          <button class="folder-card-action" data-action="menu" title="Tùy chọn">
            <i class="fa-solid fa-ellipsis-vertical"></i>
          </button>
        </div>
      `;
    }).join('');

    // Events
    grid.querySelectorAll('.folder-card').forEach(card => {
      const folderId = card.dataset.folderId;
      card.addEventListener('click', (e) => {
        if (e.target.closest('.folder-card-action')) return;
        this.currentFolderId = folderId;
        this.renderExplorer();
      });

      const menuBtn = card.querySelector('.folder-card-action');
      if (menuBtn) {
        menuBtn.addEventListener('click', (e) => {
          e.stopPropagation();
          this.showContextMenu(e, 'folder', folderId);
        });
      }
    });
  }

  renderDocuments() {
    const grid = document.getElementById('files-grid');
    const listTable = document.getElementById('files-list-container');
    const section = document.getElementById('files-section');
    const countEl = document.getElementById('files-count');
    const emptyState = document.getElementById('explorer-empty-state');
    if (!grid || !section) return;

    // Filter documents
    let currentDocs = this.documents.filter(d => d.folderId === this.currentFolderId);
    if (this.searchQuery) {
      const q = this.searchQuery.toLowerCase();
      currentDocs = this.documents.filter(d => 
        d.title.toLowerCase().includes(q) || 
        (d.preview && d.preview.toLowerCase().includes(q))
      );
    }

    // Sort
    if (this.sortBy === 'newest') {
      currentDocs.sort((a, b) => b.updatedAt - a.updatedAt);
    } else if (this.sortBy === 'name') {
      currentDocs.sort((a, b) => a.title.localeCompare(b.title));
    }

    if (countEl) countEl.textContent = `(${currentDocs.length})`;

    const totalFoldersInView = this.folders.filter(f => f.parentId === this.currentFolderId).length;
    if (currentDocs.length === 0 && totalFoldersInView === 0) {
      section.style.display = 'none';
      if (emptyState) emptyState.style.display = 'block';
      return;
    } else {
      if (emptyState) emptyState.style.display = 'none';
      section.style.display = 'block';
    }

    if (this.viewMode === 'list') {
      grid.style.display = 'none';
      listTable.style.display = 'block';
      listTable.innerHTML = `
        <table class="files-list-table">
          <thead>
            <tr>
              <th>Tên tài liệu</th>
              <th>Ngày sửa đổi</th>
              <th>Số từ</th>
              <th style="width: 48px;"></th>
            </tr>
          </thead>
          <tbody>
            ${currentDocs.map(d => `
              <tr data-doc-id="${d.id}">
                <td>
                  <div class="list-name-cell">
                    <span class="doc-word-icon" style="width: 22px; height: 22px; font-size: 11px;">W</span>
                    <span>${this.escapeHtml(d.title)}</span>
                  </div>
                </td>
                <td style="color: var(--text-secondary); font-size: 12px;">${this.formatDate(d.updatedAt)}</td>
                <td style="color: var(--text-secondary); font-size: 12px;">${(d.wordCount || 0).toLocaleString()} từ</td>
                <td>
                  <button class="doc-card-menu-btn" data-action="menu">
                    <i class="fa-solid fa-ellipsis-vertical"></i>
                  </button>
                </td>
              </tr>
            `).join('')}
          </tbody>
        </table>
      `;

      listTable.querySelectorAll('tbody tr').forEach(row => {
        const docId = row.dataset.docId;
        row.addEventListener('click', (e) => {
          if (e.target.closest('[data-action="menu"]')) return;
          this.openDocument(docId);
        });

        const menuBtn = row.querySelector('[data-action="menu"]');
        if (menuBtn) {
          menuBtn.addEventListener('click', (e) => {
            e.stopPropagation();
            this.showContextMenu(e, 'document', docId);
          });
        }
      });
    } else {
      // Grid view
      listTable.style.display = 'none';
      grid.style.display = 'grid';

      grid.innerHTML = currentDocs.map(d => `
        <div class="doc-card" data-doc-id="${d.id}">
          <div class="doc-card-header">
            <div class="doc-badge">
              <span class="doc-word-icon">W</span>
              <span class="doc-format-tag">DOCX</span>
            </div>
            <button class="doc-card-menu-btn" data-action="menu" title="Tùy chọn">
              <i class="fa-solid fa-ellipsis-vertical"></i>
            </button>
          </div>
          
          <div class="doc-card-preview">
            ${this.escapeHtml(d.preview || 'Tài liệu chưa có nội dung...')}
          </div>

          <div class="doc-card-body">
            <div class="doc-card-title" title="${this.escapeHtml(d.title)}">
              ${this.escapeHtml(d.title)}
            </div>
            <div class="doc-card-meta">
              <span class="doc-meta-item">
                <i class="fa-regular fa-clock"></i> ${this.formatDate(d.updatedAt)}
              </span>
              <span class="doc-meta-item">
                <i class="fa-solid fa-pen-nib"></i> ${(d.wordCount || 0).toLocaleString()} từ
              </span>
            </div>
          </div>
        </div>
      `).join('');

      // Events
      grid.querySelectorAll('.doc-card').forEach(card => {
        const docId = card.dataset.docId;
        card.addEventListener('click', (e) => {
          if (e.target.closest('.doc-card-menu-btn')) return;
          this.openDocument(docId);
        });

        const menuBtn = card.querySelector('.doc-card-menu-btn');
        if (menuBtn) {
          menuBtn.addEventListener('click', (e) => {
            e.stopPropagation();
            this.showContextMenu(e, 'document', docId);
          });
        }
      });
    }
  }

  // =========================================================================
  // 5. CONTEXT MENU & MODALS HANDLING
  // =========================================================================
  showContextMenu(event, type, targetId) {
    const menu = document.getElementById('explorer-context-menu');
    if (!menu) return;

    if (type === 'folder') {
      const folder = this.folders.find(f => f.id === targetId);
      menu.innerHTML = `
        <button class="context-menu-item" id="cmenu-open-folder">
          <i class="fa-solid fa-folder-open text-blue"></i> Mở thư mục
        </button>
        <button class="context-menu-item" id="cmenu-rename-folder">
          <i class="fa-solid fa-pen-to-square"></i> Đổi tên
        </button>
        <div class="context-menu-divider"></div>
        <button class="context-menu-item danger" id="cmenu-delete-folder">
          <i class="fa-solid fa-trash-can"></i> Xóa thư mục
        </button>
      `;

      document.getElementById('cmenu-open-folder').onclick = () => {
        this.currentFolderId = targetId;
        this.renderExplorer();
        this.hideContextMenu();
      };
      document.getElementById('cmenu-rename-folder').onclick = () => {
        this.hideContextMenu();
        this.showRenameModal('folder', targetId, folder ? folder.name : '');
      };
      document.getElementById('cmenu-delete-folder').onclick = () => {
        this.hideContextMenu();
        this.deleteFolder(targetId);
      };
    } else {
      const doc = this.documents.find(d => d.id === targetId);
      menu.innerHTML = `
        <button class="context-menu-item" id="cmenu-open-doc">
          <i class="fa-solid fa-pen-nib text-blue"></i> Mở soạn thảo
        </button>
        <button class="context-menu-item" id="cmenu-download-doc">
          <i class="fa-solid fa-download"></i> Tải về (.docx)
        </button>
        <button class="context-menu-item" id="cmenu-rename-doc">
          <i class="fa-solid fa-pen-to-square"></i> Đổi tên
        </button>
        <button class="context-menu-item" id="cmenu-duplicate-doc">
          <i class="fa-regular fa-copy"></i> Nhân bản
        </button>
        <button class="context-menu-item" id="cmenu-move-doc">
          <i class="fa-solid fa-arrow-right-to-bracket"></i> Di chuyển vào thư mục
        </button>
        <div class="context-menu-divider"></div>
        <button class="context-menu-item danger" id="cmenu-delete-doc">
          <i class="fa-solid fa-trash-can"></i> Xóa tài liệu
        </button>
      `;

      document.getElementById('cmenu-open-doc').onclick = () => {
        this.openDocument(targetId);
        this.hideContextMenu();
      };
      document.getElementById('cmenu-download-doc').onclick = () => {
        this.downloadDocx(targetId);
        this.hideContextMenu();
      };
      document.getElementById('cmenu-rename-doc').onclick = () => {
        this.hideContextMenu();
        this.showRenameModal('document', targetId, doc ? doc.title : '');
      };
      document.getElementById('cmenu-duplicate-doc').onclick = () => {
        this.duplicateDocument(targetId);
        this.hideContextMenu();
      };
      document.getElementById('cmenu-move-doc').onclick = () => {
        this.showMoveModal(targetId);
        this.hideContextMenu();
      };
      document.getElementById('cmenu-delete-doc').onclick = () => {
        this.deleteDocument(targetId);
        this.hideContextMenu();
      };
    }

    // Position menu
    menu.style.display = 'block';
    const x = Math.min(event.clientX, window.innerWidth - 220);
    const y = Math.min(event.clientY, window.innerHeight - 250);
    menu.style.left = `${x}px`;
    menu.style.top = `${y}px`;
  }

  hideContextMenu() {
    const menu = document.getElementById('explorer-context-menu');
    if (menu) menu.style.display = 'none';
  }

  showRenameModal(type, id, currentName) {
    const modal = document.getElementById('modal-rename');
    const input = document.getElementById('input-rename-value');
    const title = document.getElementById('modal-rename-title');
    if (!modal || !input) return;

    if (title) title.textContent = type === 'folder' ? 'Đổi tên thư mục' : 'Đổi tên tài liệu';
    input.value = currentName;
    modal.classList.add('show');
    input.focus();
    input.select();

    const saveBtn = document.getElementById('btn-confirm-rename');
    saveBtn.onclick = () => {
      const val = input.value.trim();
      if (val) {
        if (type === 'folder') this.renameFolder(id, val);
        else this.renameDocument(id, val);
        modal.classList.remove('show');
      }
    };
  }

  showMoveModal(docId) {
    const modal = document.getElementById('modal-move');
    const select = document.getElementById('select-move-destination');
    if (!modal || !select) return;

    const doc = this.documents.find(d => d.id === docId);
    if (!doc) return;

    let optionsHtml = `<option value="root">📁 [Thư mục gốc - Tất cả tệp]</option>`;
    this.folders.forEach(f => {
      optionsHtml += `<option value="${f.id}" ${doc.folderId === f.id ? 'selected' : ''}>📂 ${this.escapeHtml(f.name)}</option>`;
    });

    select.innerHTML = optionsHtml;
    modal.classList.add('show');

    const saveBtn = document.getElementById('btn-confirm-move');
    saveBtn.onclick = () => {
      const target = select.value === 'root' ? null : select.value;
      this.moveDocument(docId, target);
      modal.classList.remove('show');
    };
  }

  showNewFolderModal() {
    const modal = document.getElementById('modal-new-folder');
    const input = document.getElementById('input-new-folder-name');
    if (!modal || !input) return;

    input.value = 'Thư mục mới';
    modal.classList.add('show');
    input.focus();
    input.select();

    let selectedColor = '#185abd';
    const colorBtns = modal.querySelectorAll('.color-option-btn');
    colorBtns.forEach(btn => {
      btn.onclick = () => {
        colorBtns.forEach(b => b.classList.remove('selected'));
        btn.classList.add('selected');
        selectedColor = btn.dataset.color;
      };
    });

    const createBtn = document.getElementById('btn-confirm-new-folder');
    createBtn.onclick = () => {
      const val = input.value.trim();
      if (val) {
        this.createFolder(val, selectedColor);
        modal.classList.remove('show');
      }
    };
  }

  showAccountModal() {
    const modal = document.getElementById('modal-account');
    const input = document.getElementById('input-sync-account-name');
    if (!modal || !input) return;

    input.value = this.activeUser;
    modal.classList.add('show');

    const switchBtn = document.getElementById('btn-switch-account');
    switchBtn.onclick = () => {
      const val = input.value.trim().toLowerCase().replace(/[^a-z0-9_-]/g, '');
      if (val) {
        this.activeUser = val;
        localStorage.setItem(this.userKey, val);
        modal.classList.remove('show');
        this.loadLocalData();
        this.syncFromCloud();
        this.showToast(`Đã chuyển sang tài khoản "${val}"`);
      }
    };
  }

  showQrSyncModal() {
    const modal = document.getElementById('modal-qr-sync');
    if (!modal) return;

    const isFileProto = window.location.protocol === 'file:';
    const warningEl = document.getElementById('file-protocol-warning');
    if (warningEl) {
      warningEl.style.display = isFileProto ? 'block' : 'none';
    }

    // Default network mode: if on http/https, use online; if on file:, default to wifi
    const initialMode = isFileProto ? 'wifi' : 'online';
    if (window.switchQrNetworkMode) {
      window.switchQrNetworkMode(initialMode);
    }

    modal.classList.add('show');
  }

  // =========================================================================
  // 6. EXPORT & IMPORT WORD DOCX
  // =========================================================================
  downloadDocx(docId) {
    const doc = this.documents.find(d => d.id === docId);
    if (!doc) return;

    const fileName = doc.title.endsWith('.docx') ? doc.title : doc.title + '.docx';
    
    // Create standard Microsoft Word XML/HTML document bundle that MS Word / WPS natively opens as DOCX
    const wordTemplate = `
      <html xmlns:o='urn:schemas-microsoft-com:office:office' 
            xmlns:w='urn:schemas-microsoft-com:office:word' 
            xmlns='http://www.w3.org/TR/REC-html40'>
      <head>
        <meta charset='utf-8'>
        <title>${this.escapeHtml(doc.title)}</title>
        <!--[if gte mso 9]>
        <xml>
          <w:WordDocument>
            <w:View>Print</w:View>
            <w:Zoom>100</w:Zoom>
            <w:DoNotOptimizeForBrowser/>
          </w:WordDocument>
        </xml>
        <![endif]-->
        <style>
          @page {
            size: 21cm 29.7cm;
            margin: 2.54cm 2.54cm 2.54cm 2.54cm;
            mso-page-orientation: portrait;
          }
          body {
            font-family: 'Calibri', 'Segoe UI', Arial, sans-serif;
            font-size: 11pt;
            line-height: 1.5;
            color: #242424;
          }
          h1 { font-size: 22pt; color: #185abd; font-weight: bold; }
          h2 { font-size: 16pt; color: #103f82; font-weight: bold; }
          h3 { font-size: 13pt; color: #333333; font-weight: bold; }
          table { border-collapse: collapse; width: 100%; margin: 12pt 0; }
          td, th { border: 1px solid #999; padding: 6pt 8pt; }
          img { max-width: 100%; height: auto; }
        </style>
      </head>
      <body>
        ${doc.content}
      </body>
      </html>
    `;

    const blob = new Blob([wordTemplate], {
      type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document;charset=utf-8'
    });
    
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => {
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    }, 150);

    this.showToast(`Đã xuất tệp "${fileName}"!`);
  }

  handleDocxUpload(file) {
    if (!file) return;
    const fileName = file.name;

    this.showToast(`Đang đọc tệp "${fileName}"...`, 'fa-solid fa-spinner fa-spin');

    const reader = new FileReader();
    reader.onload = (e) => {
      const arrayBuffer = e.target.result;
      
      // If mammoth.js is present
      if (window.mammoth) {
        window.mammoth.convertToHtml({ arrayBuffer: arrayBuffer })
          .then(result => {
            const html = result.value || `<p>Tài liệu tải lên từ ${fileName}</p>`;
            this.createDocument(fileName, html);
          })
          .catch(err => {
            console.warn('Mammoth parse error, falling back:', err);
            this.createDocument(fileName, `<p>Nội dung tải lên từ ${fileName}</p>`);
          });
      } else {
        // Fallback: create document with file name
        this.createDocument(fileName, `<h1 style="color: #185abd;">${fileName}</h1><p>Tài liệu được tải lên thành công.</p>`);
      }
    };
    reader.readAsArrayBuffer(file);
  }

  // =========================================================================
  // 7. VIEW SWITCHER (EXPLORER <-> WORD EDITOR)
  // =========================================================================
  switchView(targetView) {
    const explorerEl = document.getElementById('view-explorer');
    const editorEl = document.getElementById('view-editor');

    if (targetView === 'explorer') {
      if (editorEl) editorEl.classList.remove('active');
      if (explorerEl) explorerEl.classList.add('active');
      this.renderExplorer();
      window.scrollTo(0, 0);
    } else {
      if (explorerEl) explorerEl.classList.remove('active');
      if (editorEl) editorEl.classList.add('active');
    }
  }

  // =========================================================================
  // 8. EVENT BINDINGS
  // =========================================================================
  initEventBindings() {
    // New dropdown toggle
    const btnNew = document.getElementById('btn-command-new');
    const dropdownNew = document.getElementById('command-new-dropdown');
    if (btnNew && dropdownNew) {
      btnNew.addEventListener('click', (e) => {
        e.stopPropagation();
        dropdownNew.classList.toggle('show');
      });
      document.addEventListener('click', () => dropdownNew.classList.remove('show'));
    }

    // New Document action
    const btnNewDoc = document.getElementById('btn-new-document');
    if (btnNewDoc) {
      btnNewDoc.addEventListener('click', () => {
        dropdownNew.classList.remove('show');
        this.createDocument('Tài liệu ' + (this.documents.length + 1) + '.docx');
      });
    }

    // New Folder action
    const btnNewFolder = document.getElementById('btn-new-folder');
    if (btnNewFolder) {
      btnNewFolder.addEventListener('click', () => {
        dropdownNew.classList.remove('show');
        this.showNewFolderModal();
      });
    }

    // Upload Docx file
    const fileUploadInput = document.getElementById('input-upload-docx');
    const btnUpload = document.getElementById('btn-command-upload');
    if (btnUpload && fileUploadInput) {
      btnUpload.addEventListener('click', () => fileUploadInput.click());
      fileUploadInput.addEventListener('change', (e) => {
        if (e.target.files && e.target.files[0]) {
          this.handleDocxUpload(e.target.files[0]);
          e.target.value = '';
        }
      });
    }

    // Search Box
    const searchInput = document.getElementById('explorer-search-input');
    const clearSearchBtn = document.getElementById('btn-clear-search');
    if (searchInput) {
      searchInput.addEventListener('input', (e) => {
        this.searchQuery = e.target.value.trim();
        if (clearSearchBtn) {
          clearSearchBtn.style.display = this.searchQuery ? 'block' : 'none';
        }
        this.renderFolders();
        this.renderDocuments();
      });
    }
    if (clearSearchBtn) {
      clearSearchBtn.addEventListener('click', () => {
        if (searchInput) searchInput.value = '';
        this.searchQuery = '';
        clearSearchBtn.style.display = 'none';
        this.renderFolders();
        this.renderDocuments();
      });
    }

    // View Switch buttons (Grid / List)
    const btnViewGrid = document.getElementById('btn-view-grid');
    const btnViewList = document.getElementById('btn-view-list');
    if (btnViewGrid && btnViewList) {
      btnViewGrid.addEventListener('click', () => {
        this.viewMode = 'grid';
        localStorage.setItem('clouddocs_view_mode', 'grid');
        btnViewGrid.classList.add('active');
        btnViewList.classList.remove('active');
        this.renderDocuments();
      });
      btnViewList.addEventListener('click', () => {
        this.viewMode = 'list';
        localStorage.setItem('clouddocs_view_mode', 'list');
        btnViewList.classList.add('active');
        btnViewGrid.classList.remove('active');
        this.renderDocuments();
      });
    }

    // Mobile FAB button
    const fabBtn = document.getElementById('fab-new-btn');
    if (fabBtn) {
      fabBtn.addEventListener('click', () => {
        this.createDocument('Tài liệu di động ' + (this.documents.length + 1) + '.docx');
      });
    }

    // Link Mobile / QR modal
    const btnLinkMobile = document.getElementById('btn-link-mobile');
    if (btnLinkMobile) {
      btnLinkMobile.addEventListener('click', () => this.showQrSyncModal());
    }

    // Account pill
    const userBtn = document.getElementById('user-account-btn');
    if (userBtn) {
      userBtn.addEventListener('click', () => this.showAccountModal());
    }

    // Back to Drive navigation button inside Editor
    const btnBackToDrive = document.getElementById('btn-back-to-drive');
    if (btnBackToDrive) {
      btnBackToDrive.addEventListener('click', () => {
        this.saveActiveDocument(true);
        this.switchView('explorer');
      });
    }

    // Editor AutoSave sync hooking
    const docContent = document.getElementById('document-content');
    const docTitleInput = document.getElementById('doc-title-input');
    if (docContent) {
      docContent.addEventListener('input', () => {
        this.saveActiveDocument(false);
      });
    }
    if (docTitleInput) {
      docTitleInput.addEventListener('input', () => {
        this.saveActiveDocument(false);
      });
    }

    // Close Modals on click outside or close button
    document.querySelectorAll('.modal-backdrop-custom').forEach(modal => {
      modal.addEventListener('click', (e) => {
        if (e.target === modal || e.target.closest('.modal-close-btn') || e.target.closest('.btn-cancel-modal')) {
          modal.classList.remove('show');
        }
      });
    });

    // Close Context Menu on click outside
    document.addEventListener('click', () => this.hideContextMenu());

    // Copy Sync URL button in QR modal
    const btnCopySyncLink = document.getElementById('btn-copy-sync-link');
    if (btnCopySyncLink) {
      btnCopySyncLink.addEventListener('click', () => {
        const input = document.getElementById('input-sync-share-url');
        if (input) {
          input.select();
          navigator.clipboard.writeText(input.value)
            .then(() => this.showToast('Đã sao chép link đồng bộ! Gửi qua Zalo/Telegram để mở trên ĐT.'))
            .catch(() => this.showToast('Đã chọn link, hãy bấm sao chép.'));
        }
      });
    }

    // Sync badge click triggers immediate refresh
    const syncBadge = document.getElementById('explorer-sync-badge');
    if (syncBadge) {
      syncBadge.addEventListener('click', () => {
        this.showToast('Đang kiểm tra dữ liệu đồng bộ...', 'fa-solid fa-arrows-rotate fa-spin');
        this.syncFromCloud();
      });
    }
    // PWA Install Button
    const btnInstallPwa = document.getElementById('btn-install-pwa');
    if (btnInstallPwa) {
      btnInstallPwa.addEventListener('click', () => {
        if (window.deferredInstallPrompt) {
          window.deferredInstallPrompt.prompt();
          window.deferredInstallPrompt.userChoice.then((choice) => {
            if (choice.outcome === 'accepted') {
              this.showToast('Đang cài đặt CloudDocs vào thiết bị...', 'fa-solid fa-download');
            }
            window.deferredInstallPrompt = null;
          });
        } else {
          this.showInstallGuideModal();
        }
      });
    }

    // Trigger PWA install inside modal
    const btnTriggerInstall = document.getElementById('btn-trigger-pwa-install');
    if (btnTriggerInstall) {
      btnTriggerInstall.addEventListener('click', () => {
        if (window.deferredInstallPrompt) {
          window.deferredInstallPrompt.prompt();
          window.deferredInstallPrompt.userChoice.then(() => {
            window.deferredInstallPrompt = null;
          });
        } else {
          this.showToast('Hãy mở menu ⋮ của trình duyệt và chọn "Cài đặt ứng dụng" hoặc "Thêm vào MH chính"');
        }
      });
    }

    // Dynamic QR update on URL edit
    const inputSyncShareUrl = document.getElementById('input-sync-share-url');
    if (inputSyncShareUrl) {
      inputSyncShareUrl.addEventListener('input', (e) => {
        const val = e.target.value.trim();
        if (val && window.renderQrCode) {
          window.renderQrCode(val);
          if (val.startsWith('http')) {
            localStorage.setItem('clouddocs_github_url', val);
          }
        }
      });
    }
  }

  showInstallGuideModal() {
    const modal = document.getElementById('modal-install-guide');
    if (modal) {
      // Auto-detect iOS to pre-select iOS tab
      const isIos = /iPad|iPhone|iPod/.test(navigator.userAgent) && !window.MSStream;
      if (window.switchPwaTab) {
        window.switchPwaTab(isIos ? 'ios' : 'android');
      }
      modal.classList.add('show');
    }
  }

  // =========================================================================
  // 9. UTILITIES
  // =========================================================================
  formatDate(ts) {
    if (!ts) return 'Vừa xong';
    const now = Date.now();
    const diff = Math.floor((now - ts) / 1000); // seconds
    if (diff < 60) return 'Vừa xong';
    if (diff < 3600) return `${Math.floor(diff / 60)} phút trước`;
    if (diff < 86400) return `${Math.floor(diff / 3600)} giờ trước`;
    const d = new Date(ts);
    return `${d.getDate().toString().padStart(2, '0')}/${(d.getMonth() + 1).toString().padStart(2, '0')}/${d.getFullYear()}`;
  }

  escapeHtml(str) {
    if (!str) return '';
    return str
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');
  }

  showToast(msg, icon = 'fa-solid fa-circle-check') {
    const container = document.getElementById('toast-container') || document.body;
    const toast = document.createElement('div');
    toast.className = 'office-toast';
    toast.style.cssText = 'position: fixed; bottom: 24px; right: 24px; background: #242424; color: #fff; padding: 10px 18px; border-radius: 6px; display: flex; align-items: center; gap: 10px; font-size: 13px; z-index: 999999; box-shadow: 0 4px 16px rgba(0,0,0,0.25); animation: fadeInUp 0.25s ease-out;';
    toast.innerHTML = `<i class="${icon}" style="color: #4ade80;"></i> <span>${msg}</span>`;
    container.appendChild(toast);

    setTimeout(() => {
      toast.style.opacity = '0';
      toast.style.transform = 'translateY(8px)';
      toast.style.transition = 'all 0.3s ease';
      setTimeout(() => toast.remove(), 300);
    }, 2800);
  }
}

// Global PWA Tab Switcher
window.switchPwaTab = function(platform) {
  document.querySelectorAll('.pwa-tab-btn').forEach(btn => btn.classList.remove('active'));
  document.querySelectorAll('.pwa-tab-content').forEach(c => c.style.display = 'none');

  const activeBtn = document.getElementById(`tab-btn-${platform}`);
  const activeContent = document.getElementById(`pwa-content-${platform}`);

  if (activeBtn) activeBtn.classList.add('active');
  if (activeContent) activeContent.style.display = 'block';
};

// Global QR Network Mode Switcher & Renderer
window.renderQrCode = function(url) {
  const qrBox = document.getElementById('qr-code-box');
  const linkInput = document.getElementById('input-sync-share-url');
  if (linkInput) linkInput.value = url;
  if (!qrBox) return;

  qrBox.innerHTML = '';
  if (window.QRCode) {
    new window.QRCode(qrBox, {
      text: url,
      width: 180,
      height: 180,
      colorDark: '#185abd',
      colorLight: '#ffffff',
      correctLevel: window.QRCode.CorrectLevel.M
    });
  } else {
    qrBox.innerHTML = `<p style="padding: 10px; font-size: 12px; color: #666;">Dùng liên kết bên dưới để mở trên điện thoại.</p>`;
  }
};

window.switchQrNetworkMode = function(mode) {
  const tabWifi = document.getElementById('qr-tab-wifi');
  const tabOnline = document.getElementById('qr-tab-online');
  const instruction = document.getElementById('qr-instruction-text');
  const hint = document.getElementById('hint-qr-url');
  const activeUser = window.cloudDocs ? window.cloudDocs.activeUser : 'user';

  if (tabWifi && tabOnline) {
    tabWifi.classList.toggle('active', mode === 'wifi');
    tabOnline.classList.toggle('active', mode === 'online');
  }

  let targetUrl = '';
  if (mode === 'wifi') {
    // Local IP address of user's machine on port 5502
    targetUrl = `http://172.68.1.47:5502/index.html?user=${activeUser}`;
    if (instruction) {
      instruction.innerHTML = 'Đảm bảo điện thoại và máy tính kết nối <strong>cùng 1 mạng Wi-Fi</strong>, sau đó quét mã QR:';
    }
    if (hint) {
      hint.innerHTML = 'IP Wi-Fi máy tính của bạn: <code>http://172.68.1.47:5502</code>';
    }
  } else {
    // Online mode (GitHub Pages or public HTTPS)
    const savedGithubUrl = localStorage.getItem('clouddocs_github_url') || 'https://<tên_github>.github.io/word/';
    if (window.location.protocol.startsWith('http')) {
      const cur = new URL(window.location.href);
      cur.searchParams.set('user', activeUser);
      targetUrl = cur.toString();
    } else {
      targetUrl = `${savedGithubUrl.replace(/\/$/, '')}/index.html?user=${activeUser}`;
    }
    if (instruction) {
      instruction.innerHTML = 'Quét mã để mở qua Internet/4G (sau khi đã đưa lên <strong>GitHub Pages</strong>):';
    }
    if (hint) {
      hint.innerHTML = 'Bạn có thể chỉnh sửa link GitHub Pages ở ô trên và mã QR sẽ tự động cập nhật.';
    }
  }

  window.renderQrCode(targetUrl);
};

// Global PWA Install Prompt Listener
window.deferredInstallPrompt = null;
window.addEventListener('beforeinstallprompt', (e) => {
  e.preventDefault();
  window.deferredInstallPrompt = e;
  console.log('[PWA] beforeinstallprompt captured');
  const btn = document.getElementById('btn-install-pwa');
  if (btn) btn.classList.add('pwa-btn-highlight');
});

window.addEventListener('appinstalled', () => {
  window.deferredInstallPrompt = null;
  console.log('[PWA] Application successfully installed to home screen!');
  if (window.cloudDocs) {
    window.cloudDocs.showToast('Đã cài đặt CloudDocs thành công vào thiết bị!');
  }
});

// Register Service Worker for PWA Offline & Mobile Independence
if ('serviceWorker' in navigator) {
  window.addEventListener('load', () => {
    navigator.serviceWorker.register('./sw.js', { scope: './' })
      .then((reg) => {
        console.log('[PWA] Service Worker registered successfully:', reg.scope);
      })
      .catch((err) => {
        console.log('[PWA] Service Worker registration failed:', err);
      });
  });
}

// Instantiate engine when DOM is ready
window.addEventListener('DOMContentLoaded', () => {
  window.cloudDocs = new CloudDocsEngine();
});

