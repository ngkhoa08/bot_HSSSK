(function() {
    'use strict';

    // --- BẢN VÁ LỖI QUAN TRỌNG: LÀM SẠCH BỘ NHỚ KHI CHẠY LẠI BOT ---
    if (window.hsskBotInterval) {
        clearInterval(window.hsskBotInterval);
    }
    const oldPanel = document.getElementById('hssk-bot-panel');
    if (oldPanel) {
        oldPanel.remove();
    }
    // ---------------------------------------------------------------

    // Biến toàn cục lưu trữ trạng thái
    window.hsskBotQueue = [];
    window.totalChuanHoa = 0; 
    window.currentSearchId = null;
    
    window.themMoiQueue = [];
    window.totalThemMoi = 0;
    window.isThemMoiRunning = false;
    window.currentThemMoiId = null; 

    window.ghephoQueue = [];
    window.totalGhephoFamilies = 0;
    window.isGhephoRunning = false;
    window.currentGhephoId = null;

    // Các biến phụ trợ
    window.lastCollectedMedicalId = ""; 
    window.isCompactMode = true;

    function calculateAge(dobString) {
        if (!dobString) return null;
        const parts = dobString.split('/');
        if (parts.length !== 3) return null;
        const dob = new Date(parts[2], parts[1] - 1, parts[0]);
        const today = new Date();
        let age = today.getFullYear() - dob.getFullYear();
        const m = today.getMonth() - dob.getMonth();
        if (m < 0 || (m === 0 && today.getDate() < dob.getDate())) age--;
        return age;
    }

    window.textChuanHoa = window.textChuanHoa || "";
    window.textGhepHo = window.textGhepHo || "";
    window.textThemMoi = window.textThemMoi || "";
    window.textDienNhanhXa = window.textDienNhanhXa || "Xã Quốc Oai";
    window.textDienNhanhThon = window.textDienNhanhThon || "hoa vôi";
    window.currentTab = window.currentTab || "chuanhoa";

    async function switchToTab(tabText, hashFallback) {
        let tabFound = false;
        const tabs = document.querySelectorAll('.label-tab span');
        for (let t of tabs) {
            if (t.innerText.trim() === tabText) {
                const tabLabel = t.closest('.mat-tab-label');
                if (tabLabel) {
                    tabLabel.click();
                    tabFound = true;
                    break;
                }
            }
        }
        if (!tabFound) {
            const menuLinks = document.querySelectorAll('a.matero-sidemenu-link');
            for (let a of menuLinks) {
                if (a.href.includes(hashFallback)) {
                    a.click();
                    tabFound = true;
                    break;
                }
            }
        }
        if (!tabFound) window.location.hash = hashFallback;
        await sleep(1000);
    }

    function createBotPanel() {
        if (document.getElementById('hssk-bot-panel')) return;

        const panel = document.createElement('div');
        panel.id = 'hssk-bot-panel';
        panel.style.cssText = `
            position: fixed; top: 70px; right: 20px; z-index: 999999; 
            background: #ffffff; border: 2px solid #28a745; padding: 0; 
            border-radius: 8px; box-shadow: 0 10px 25px rgba(0,0,0,0.25); 
            width: 360px; font-family: 'Inter', sans-serif; transition: opacity 0.3s;
        `;

        panel.innerHTML = `
            <!-- HEADER -->
            <div id="hssk_drag_header" style="display: flex; justify-content: space-between; align-items: center; background: #e8f5e9; padding: 10px 15px; border-bottom: 1px solid #c3e6cb; border-radius: 6px 6px 0 0; cursor: move; user-select: none;">
                <h3 style="margin:0; color:#28a745; font-size: 15px; font-weight: bold; pointer-events: none;">🤖 Bot HSSK v10.0 (Smart Group)</h3>
                <button id="hssk_toggle" style="background: none; border: none; cursor: pointer; font-size: 16px; outline: none; padding: 0;">➖</button>
            </div>
            
            <div id="hssk_body" style="max-height: 85vh; overflow-y: auto; overflow-x: hidden; padding: 10px;">
                <!-- TABS -->
                <div style="display: flex; border-bottom: 2px solid #ddd; margin-bottom: 12px;">
                    <button id="tab_chuanhoa" style="flex: 1; padding: 8px 2px; border: none; background: none; font-weight: bold; color: #28a745; border-bottom: 3px solid #28a745; cursor: pointer; transition: 0.3s; font-size: 12px;">🔍 CHUẨN HÓA</button>
                    <button id="tab_ghepho" style="flex: 1; padding: 8px 2px; border: none; background: none; font-weight: bold; color: #666; border-bottom: 3px solid transparent; cursor: pointer; transition: 0.3s; font-size: 12px;">👨‍👩‍👧‍👦 GHÉP HỘ</button>
                    <button id="tab_themmoi" style="flex: 1; padding: 8px 2px; border: none; background: none; font-weight: bold; color: #666; border-bottom: 3px solid transparent; cursor: pointer; transition: 0.3s; font-size: 12px;">➕ THÊM MỚI</button>
                </div>

                <!-- TAB 1: CHUẨN HÓA -->
                <div id="content_chuanhoa" style="display: none;">
                    <!-- HƯỚNG DẪN NẠP DỮ LIỆU -->
                    <div style="margin-bottom: 10px;">
                        <button id="btn_huong_dan" style="width: 100%; background: #ffc107; color: #333; border: 1px solid #ffb300; padding: 8px; font-weight: bold; border-radius: 4px; font-size: 12px; cursor: pointer; box-shadow: 0 2px 4px rgba(0,0,0,0.1);">ℹ️ HƯỚNG DẪN COPY EXCEL (CHUẨN HÓA)</button>
                    </div>

                    <div style="background: #f8f9fa; padding: 8px; border-radius: 6px; border: 1px dashed #ccc; margin-bottom: 10px;">
                        <textarea id="hssk_excel_input" placeholder="Copy cả dòng từ Excel dán vào đây..." style="width: 100%; height: 55px; margin-bottom: 5px; padding: 5px; border: 1px solid #ccc; border-radius: 4px; font-size: 11px;"></textarea>
                        <div style="display: flex; gap: 5px;">
                            <button id="btn_import" style="width: 100%; background: #17a2b8; color: white; border: none; padding: 6px; font-weight: bold; border-radius: 4px; font-size: 11px;">📥 NẠP DỮ LIỆU</button>
                        </div>
                    </div>

                    <div style="margin-bottom: 10px; border-left: 3px solid #ff9800; padding-left: 8px;">
                        <div style="font-size: 12px; font-weight: bold; color: #333;">Đang chờ: <span id="hssk_queue_count" style="color: #d32f2f;">0</span> / Tổng nạp: <span id="hssk_total_count">0</span></div>
                        <div style="font-size: 12px; font-weight: bold; color: #0056b3; margin-top: 4px;">Đang xử lý: <br><span id="hssk_current_person" style="font-size: 13px; font-weight: normal; color: #333;">Chưa có</span></div>
                        <div id="hssk_status" style="font-size: 11px; color: #666; margin-top: 6px; font-style: italic;">Trạng thái: Đang nghỉ...</div>
                    </div>

                    <div style="display: flex; flex-direction: column; gap: 6px;">
                        <div style="display: flex; gap: 6px;">
                            <button id="btn_auto_start" style="flex: 2; background: #0056b3; color: white; border: none; padding: 8px; font-weight: bold; cursor: pointer; border-radius: 4px; font-size: 12px;">▶️ CHẠY AUTO</button>
                            <button id="btn_fill_manual" style="flex: 1; background: #17a2b8; color: white; border: none; padding: 8px; font-weight: bold; cursor: pointer; border-radius: 4px; font-size: 12px;">✍️ ĐIỀN TAY</button>
                        </div>
                        <div style="display: flex; gap: 6px;">
                            <button id="btn_approve" style="flex: 2; background: #28a745; color: white; border: none; padding: 8px; font-weight: bold; cursor: pointer; border-radius: 4px; font-size: 12px; display: none;">✅ DUYỆT & TIẾP TỤC</button>
                            <button id="btn_skip" style="flex: 1; background: #dc3545; color: white; border: none; padding: 8px; font-weight: bold; cursor: pointer; border-radius: 4px; font-size: 12px; display: none; transition: 0.2s;">⏭️ BỎ QUA</button>
                        </div>
                        <button id="btn_continue_after_popup" style="width: 100%; background: #6f42c1; color: white; border: none; padding: 8px; font-weight: bold; cursor: pointer; border-radius: 4px; font-size: 12px; display: none;">⏭️ ĐÃ XỬ LÝ POPUP - TIẾP TỤC</button>
                        
                        <button id="btn_auto_fill_transfer" style="width: 100%; background: #fd7e14; color: white; border: none; padding: 8px; font-weight: bold; cursor: pointer; border-radius: 4px; font-size: 12px; display: none; margin-top: 6px; box-shadow: 0 0 8px rgba(253,126,20,0.5);">🏡 TỰ ĐIỀN CHUYỂN ĐỊA BÀN</button>
                        
                        <div style="display: flex; gap: 6px; margin-top: 6px;">
                            <button id="btn_add_cccd_chuanhoa" title="Chưa có mã y tế nào được lưu" style="flex: 1; background: #e83e8c; color: white; border: none; padding: 8px; font-weight: bold; cursor: pointer; border-radius: 4px; font-size: 12px; display: block; box-shadow: 0 0 8px rgba(232,62,140,0.5);">🏷️ SỬA BẰNG MÃ Y TẾ</button>
                            <button id="btn_hd_sua_ma_chuanhoa" title="Xem hướng dẫn chức năng này" style="background: #ffc107; color: #333; border: 1px solid #ffb300; padding: 0 12px; font-weight: bold; cursor: pointer; border-radius: 4px; font-size: 14px; box-shadow: 0 2px 4px rgba(0,0,0,0.1);">❓</button>
                        </div>
                    </div>
                </div>

                <!-- TAB 2: GHÉP HỘ (GIAO DIỆN MỚI THÔNG MINH) -->
                <div id="content_ghepho" style="display: none;">
                    <div style="background: #fdfaf2; padding: 8px; border-radius: 6px; border: 1px dashed #ccc; margin-bottom: 10px;">
                        <textarea id="ghepho_excel_input" style="width: 100%; height: 65px; padding: 5px; border: 1px solid #ccc; border-radius: 4px; font-size: 11px;" placeholder="Copy CẢ KHỐI DỮ LIỆU gia đình từ Excel dán vào đây (Gồm Tên, Tên Chủ Hộ, CCCD...)..."></textarea>
                        <button id="btn_import_ghepho" style="width: 100%; margin-top:5px; background: #17a2b8; color: white; border: none; padding: 6px; font-weight: bold; border-radius: 4px; font-size: 11px;">📥 NẠP & PHÂN TÍCH DỮ LIỆU</button>
                    </div>
                    
                    <div style="margin-bottom: 10px; border-left: 3px solid #17a2b8; padding-left: 8px;">
                        <div style="font-size: 12px; font-weight: bold; color: #333;">Hộ gia đình chờ ghép: <span id="ghepho_queue_count" style="color: #d32f2f;">0</span> / <span id="ghepho_total_count">0</span></div>
                        <div style="font-size: 12px; font-weight: bold; color: #0056b3; margin-top: 4px;">Đang xử lý Hộ:<br><span id="ghepho_current_family" style="color: #17a2b8; font-size: 13px;">Chưa có</span></div>
                        <div id="ghepho_status" style="font-size: 11px; color: #666; margin-top: 4px; font-style: italic;">Trạng thái: Sẵn sàng...</div>
                    </div>

                    <div style="display: flex; flex-direction: column; gap: 6px;">
                        <div style="display: flex; gap: 6px;">
                            <button id="btn_start_ghepho" style="flex: 2; background: #28a745; color: white; border: none; padding: 8px; border-radius: 4px; font-weight: bold; cursor: pointer; font-size: 12px;">▶️ BẮT ĐẦU GHÉP</button>
                            <button id="btn_stop_ghepho" style="flex: 1; background: #dc3545; color: white; border: none; padding: 8px; border-radius: 4px; font-weight: bold; cursor: pointer; font-size: 12px;" disabled>⏸️ TẠM DỪNG</button>
                        </div>
                        <button id="btn_approve_ghepho" style="width: 100%; background: #0056b3; color: white; border: none; padding: 8px; font-weight: bold; cursor: pointer; border-radius: 4px; font-size: 12px; display: none;">✅ DUYỆT & LƯU HỘ KHẨU NÀY</button>
                        <button id="btn_skip_ghepho" style="width: 100%; background: #dc3545; color: white; border: none; padding: 8px; font-weight: bold; cursor: pointer; border-radius: 4px; font-size: 12px; display: none;">⏭️ BỎ QUA HỘ NÀY</button>
                    </div>
                </div>

                <!-- TAB 3: THÊM MỚI HÀNG LOẠT -->
                <div id="content_themmoi" style="display: none;">
                    <div style="margin-bottom: 10px;">
                        <button id="btn_huong_dan_themmoi" style="width: 100%; background: #ffc107; color: #333; border: 1px solid #ffb300; padding: 8px; font-weight: bold; border-radius: 4px; font-size: 12px; cursor: pointer; box-shadow: 0 2px 4px rgba(0,0,0,0.1);">ℹ️ HƯỚNG DẪN COPY EXCEL (THÊM MỚI)</button>
                    </div>

                    <div style="background: #fdfaf2; padding: 8px; border-radius: 6px; border: 1px dashed #ccc; margin-bottom: 10px;">
                        <textarea id="auto_themmoi_list" style="width: 100%; height: 55px; padding: 5px; border: 1px solid #ccc; border-radius: 4px; font-size: 11px;" placeholder="Copy & dán danh sách từ Excel vào đây..."></textarea>
                        <button id="btn_import_themmoi" style="width: 100%; margin-top:5px; background: #17a2b8; color: white; border: none; padding: 6px; font-weight: bold; border-radius: 4px; font-size: 11px;">📥 NẠP DỮ LIỆU</button>
                    </div>
                    
                    <div style="margin-bottom: 10px; border-left: 3px solid #6f42c1; padding-left: 8px;">
                        <div style="font-size: 12px; font-weight: bold; color: #333;">Đang chờ tạo: <span id="themmoi_queue_count" style="color: #d32f2f;">0</span> / Tổng nạp: <span id="themmoi_total_count">0</span></div>
                        <div style="font-size: 12px; font-weight: bold; color: #0056b3; margin-top: 4px;">Đang tạo mới: <br><span id="themmoi_current_person" style="color: #6f42c1; font-size: 14px; font-weight: bold;">Chưa có</span></div>
                        <div id="themmoi_status" style="font-size: 11px; color: #666; margin-top: 4px; font-style: italic;">Trạng thái: Sẵn sàng...</div>
                    </div>

                    <div style="display: flex; flex-direction: column; gap: 6px;">
                        <div style="display: flex; gap: 6px;">
                            <button id="btn_start_themmoi" style="flex: 2; background: #28a745; color: white; border: none; padding: 8px; border-radius: 4px; font-weight: bold; cursor: pointer; font-size: 12px;">▶️ BẮT ĐẦU TẠO</button>
                            <button id="btn_stop_themmoi" style="flex: 1; background: #dc3545; color: white; border: none; padding: 8px; border-radius: 4px; font-weight: bold; cursor: pointer; font-size: 12px;" disabled>⏸️ TẠM DỪNG</button>
                        </div>
                        <button id="btn_approve_themmoi" style="width: 100%; background: #0056b3; color: white; border: none; padding: 8px; font-weight: bold; cursor: pointer; border-radius: 4px; font-size: 12px; display: none;">✅ DUYỆT & LƯU</button>
                        <button id="btn_skip_themmoi" style="width: 100%; background: #dc3545; color: white; border: none; padding: 8px; font-weight: bold; cursor: pointer; border-radius: 4px; font-size: 12px; display: none;">⏭️ BỎ QUA NGƯỜI NÀY</button>
                        
                        <div style="display: flex; gap: 6px; margin-top: 5px;">
                            <button id="btn_add_cccd_themmoi" title="Chưa có mã y tế nào được lưu" style="flex: 1; background: #e83e8c; color: white; border: none; padding: 8px; font-weight: bold; cursor: pointer; border-radius: 4px; font-size: 12px; display: block; box-shadow: 0 0 8px rgba(232,62,140,0.5);">🏷️ SỬA BẰNG MÃ Y TẾ</button>
                            <button id="btn_hd_sua_ma_themmoi" title="Xem hướng dẫn chức năng này" style="background: #ffc107; color: #333; border: 1px solid #ffb300; padding: 0 12px; font-weight: bold; cursor: pointer; border-radius: 4px; font-size: 14px; box-shadow: 0 2px 4px rgba(0,0,0,0.1);">❓</button>
                        </div>
                        
                        <button id="btn_auto_fill_transfer_themmoi" style="width: 100%; background: #fd7e14; color: white; border: none; padding: 8px; font-weight: bold; cursor: pointer; border-radius: 4px; font-size: 12px; display: none; margin-top: 6px; box-shadow: 0 0 8px rgba(253,126,20,0.5);">🏡 TỰ ĐIỀN CHUYỂN ĐỊA BÀN</button>
                    </div>
                </div>

                <!-- TÍNH NĂNG GỌN MÀN HÌNH -->
                <div style="margin-top: 15px; border-top: 1px solid #eee; padding-top: 10px; text-align: center;">
                    <label style="font-size: 12px; font-weight: bold; color: #d32f2f; cursor: pointer; display: flex; align-items: center; justify-content: center; gap: 5px;">
                        <input type="checkbox" id="toggle_compact" ${window.isCompactMode ? 'checked' : ''} style="cursor: pointer; width: 15px; height: 15px;">
                        👁️ Ẩn các mục phụ (Gọn màn hình)
                    </label>
                </div>
            </div>
            <style>
                @keyframes pulse {
                    0% { transform: scale(1); box-shadow: 0 0 0 0 rgba(111, 66, 193, 0.7); }
                    50% { transform: scale(1.02); box-shadow: 0 0 0 10px rgba(111, 66, 193, 0); }
                    100% { transform: scale(1); box-shadow: 0 0 0 0 rgba(111, 66, 193, 0); }
                }
                #btn_skip:hover, #btn_skip_themmoi:hover, #btn_skip_ghepho:hover {
                    background-color: #c82333 !important;
                    transform: scale(1.02);
                }
            </style>
        `;
        document.body.appendChild(panel);

        const safeAddListener = (id, event, handler) => {
            const el = document.getElementById(id);
            if (el) el.addEventListener(event, handler);
        };

        const idsToRestore = [
            { id: 'hssk_excel_input', prop: 'textChuanHoa' },
            { id: 'ghepho_excel_input', prop: 'textGhepHo' },
            { id: 'auto_themmoi_list', prop: 'textThemMoi' }
        ];

        idsToRestore.forEach(item => {
            const el = document.getElementById(item.id);
            if (el) {
                el.value = window[item.prop];
                el.addEventListener('input', (e) => window[item.prop] = e.target.value);
            }
        });

        safeAddListener('toggle_compact', 'change', (e) => {
            window.isCompactMode = e.target.checked;
            applyCompactMode();
        });

        // SỰ KIỆN HIỂN THỊ ẢNH HƯỚNG DẪN KHI BẤM NÚT
        safeAddListener('btn_huong_dan', 'click', () => {
            let modal = document.getElementById('hssk_guide_modal');
            if (!modal) {
                modal = document.createElement('div');
                modal.id = 'hssk_guide_modal';
                modal.style.cssText = `
                    position: fixed; top: 0; left: 0; width: 100vw; height: 100vh;
                    background: rgba(0,0,0,0.85); z-index: 9999999;
                    display: flex; justify-content: center; align-items: center;
                    cursor: pointer;
                `;
                modal.innerHTML = `
                    <div style="position: relative; max-width: 90%; max-height: 90%; text-align: center;">
                        <button style="position: absolute; top: -15px; right: -15px; background: #dc3545; color: white; border: 2px solid white; border-radius: 50%; width: 32px; height: 32px; font-weight: bold; cursor: pointer; font-size: 16px;">X</button>
                        <img src="https://files.catbox.moe/5905g1.jpg" style="max-width: 100%; max-height: 90vh; border-radius: 8px; box-shadow: 0 5px 25px rgba(0,0,0,0.5);">
                    </div>
                `;
                document.body.appendChild(modal);
                modal.onclick = () => modal.style.display = 'none';
            } else {
                modal.style.display = 'flex';
            }
        });

        const header = document.getElementById('hssk_drag_header');
        let pos1 = 0, pos2 = 0, pos3 = 0, pos4 = 0;
        if (header) {
            header.onmousedown = function(e) {
                if (e.target.tagName.toLowerCase() === 'button') return; 
                e.preventDefault();
                pos3 = e.clientX;
                pos4 = e.clientY;
                document.onmouseup = closeDragElement;
                document.onmousemove = elementDrag;
            };
        }
        function elementDrag(e) {
            e.preventDefault();
            pos1 = pos3 - e.clientX;
            pos2 = pos4 - e.clientY;
            pos3 = e.clientX;
            pos4 = e.clientY;
            panel.style.top = (panel.offsetTop - pos2) + "px";
            panel.style.left = (panel.offsetLeft - pos1) + "px";
        }
        function closeDragElement() {
            document.onmouseup = null;
            document.onmousemove = null;
        }

        const switchTabUI = (tabId) => {
            window.currentTab = tabId;
            ['chuanhoa', 'ghepho', 'themmoi'].forEach(id => {
                const tabBtn = document.getElementById(`tab_${id}`);
                const contentDiv = document.getElementById(`content_${id}`);
                if (tabBtn && contentDiv) {
                    if (id === tabId) {
                        tabBtn.style.color = '#28a745';
                        tabBtn.style.borderBottomColor = '#28a745';
                        contentDiv.style.display = 'block';
                    } else {
                        tabBtn.style.color = '#666';
                        tabBtn.style.borderBottomColor = 'transparent';
                        contentDiv.style.display = 'none';
                    }
                }
            });
        };

        safeAddListener('tab_chuanhoa', 'click', () => switchTabUI('chuanhoa'));
        safeAddListener('tab_ghepho', 'click', () => switchTabUI('ghepho'));
        safeAddListener('tab_themmoi', 'click', () => switchTabUI('themmoi'));
        switchTabUI(window.currentTab); 

        let isExpanded = true;
        safeAddListener('hssk_toggle', 'click', () => {
            isExpanded = !isExpanded;
            document.getElementById('hssk_body').style.display = isExpanded ? 'block' : 'none';
            document.getElementById('hssk_toggle').innerText = isExpanded ? '➖' : '➕';
        });

        // ---------------- TAB 1: CHUẨN HÓA LISTENERS ----------------

        safeAddListener('btn_import', 'click', () => {
            const rawText = document.getElementById('hssk_excel_input').value.trim();
            if (!rawText) return alert("Vui lòng dán danh sách CCCD (hoặc Excel) vào ô trước!");
            const lines = rawText.split('\n');
            window.hsskBotQueue = [];
            lines.forEach(line => {
                let parsed = parseExcelRow(line);
                if (parsed && parsed.cccd) { window.hsskBotQueue.push(parsed); }
            });
            window.totalChuanHoa = window.hsskBotQueue.length;
            updateUI();
            document.getElementById('hssk_excel_input').value = ""; 
            window.textChuanHoa = "";
        });

        safeAddListener('btn_auto_start', 'click', async () => {
            if (window.hsskBotQueue.length === 0) return alert("Danh sách trống!");
            document.getElementById('btn_auto_start').style.display = 'none';
            document.getElementById('btn_approve').style.display = 'block';
            document.getElementById('btn_skip').style.display = 'block';
            processNextPerson();
        });

        // --- HÀM XỬ LÝ CHUNG KHI BẤM LƯU CHUẨN HÓA (TỪ BOT HOẶC TỪ WEB) ---
        async function handleApproveChuanHoa(isNativeClicked = false) {
            updateStatus(isNativeClicked ? "⏳ Phát hiện bấm Lưu trên web. Đang theo dõi..." : "⏳ Đang lưu dữ liệu...");
            setButtonsDisabled(true);
            
            if (!isNativeClicked) {
                const saveBtn = findButtonByText('Lưu');
                if (saveBtn) {
                    saveBtn.click();
                } else {
                    alert("❌ Không tìm thấy nút Lưu!");
                    setButtonsDisabled(false);
                    return;
                }
            }

            await sleep(500); // Đợi 1 chút để bắt đầu kích hoạt Loading
            let waitSave = 0;
            while (isWebLoading() && waitSave < 15000) { await sleep(300); waitSave += 300; }
            await sleep(500); 
            
            // --- BỘ QUÉT TỔNG HỢP (QUÉT CÙNG LÚC 3 TRẠNG THÁI) ---
            let isSuccess = false;
            let isDuplicate = false;
            let hasError = false;

            for (let i = 0; i < 30; i++) { // Quét liên tục trong 9 giây
                // 1. Quét tìm Popup trùng
                if (isDuplicatePopupVisible()) {
                    isDuplicate = true; break;
                }
                // 2. Quét tìm Form báo lỗi đỏ
                if (document.querySelector('.mat-error') || document.querySelector('.invalid-feedback')) {
                    hasError = true; break;
                }
                // 3. Quét tìm Toast thông báo thành công
                const toasts = document.querySelectorAll('.toast-message, .cdk-overlay-container, #toast-container, .snack-bar-container, .ngx-toastr');
                for (let toast of toasts) {
                    if (toast && toast.offsetParent !== null) {
                        const text = toast.innerText.toLowerCase();
                        if (text.includes('thành công') || text.includes('cập nhật') || text.includes('lưu') || text.includes('thêm mới')) {
                            isSuccess = true; break;
                        }
                    }
                }
                if(isSuccess) break;
                
                await sleep(300); // Chờ 0.3s rồi quét lại vòng tiếp theo
            }

            // --- XỬ LÝ DỰA TRÊN KẾT QUẢ QUÉT TÌM ĐƯỢC ---
            if (isDuplicate) {
                updateStatus("⚠️ Hệ thống báo TRÙNG LẶP. Bot đang tạm dừng!");
                document.getElementById('btn_approve').style.display = 'none';
                document.getElementById('btn_continue_after_popup').style.display = 'block';
                setButtonsDisabled(false); 
                return; 
            } else if (hasError) {
                updateStatus("⚠️ Web báo lỗi (Thiếu thông tin)! Hãy xử lý tay rồi bấm Duyệt lại.");
                setButtonsDisabled(false); 
                return; 
            } else if (isSuccess) {
                const backBtn = findButtonByText('Quay lại');
                if (backBtn) { backBtn.setAttribute('type', 'button'); backBtn.click(); }
                await sleep(300);
                let waitBack = 0;
                while (isWebLoading() && waitBack < 10000) { await sleep(300); waitBack += 300; }
                
                window.hsskBotQueue.shift();
                updateUI();
                processNextPerson();
            } else {
                updateStatus("⚠️ Không thấy phản hồi từ hệ thống! Bot tạm dừng để bạn kiểm tra.");
                setButtonsDisabled(false);
                return; 
            }
        }

        safeAddListener('btn_approve', 'click', () => {
            handleApproveChuanHoa(false);
        });

        // --- BẮT SỰ KIỆN KHI BẤM NÚT "LƯU" CỦA CHÍNH TRANG WEB ---
        if (!window.hsskGlobalSaveListenerAdded) {
            document.body.addEventListener('click', (e) => {
                let target = e.target;
                let clickedSave = false;
                
                // Truy ngược DOM để tìm xem có phải bấm vào nút Lưu không (vì đôi khi click dính vào thẻ <span> bên trong nút)
                while (target && target !== document.body) {
                    if ((target.tagName === 'BUTTON' || target.tagName === 'APP-BASE-BUTTON') && target.innerText.trim() === 'Lưu') {
                        clickedSave = true;
                        break;
                    }
                    target = target.parentElement;
                }

                if (clickedSave) {
                    // Xử lý click Lưu cho tab Chuẩn Hóa
                    if (window.currentTab === 'chuanhoa' && window.hsskBotQueue && window.hsskBotQueue.length > 0) {
                        const btnApprove = document.getElementById('btn_approve');
                        if (btnApprove && btnApprove.style.display === 'block' && !btnApprove.disabled) {
                            handleApproveChuanHoa(true);
                        }
                    }
                    // Xử lý click Lưu cho tab Ghép Hộ
                    if (window.currentTab === 'ghepho' && window.ghephoQueue && window.ghephoQueue.length > 0) {
                        const btnApproveGhepHo = document.getElementById('btn_approve_ghepho');
                        if (btnApproveGhepHo && btnApproveGhepHo.style.display === 'block' && !btnApproveGhepHo.disabled) {
                            handleApproveGhepHo(true);
                        }
                    }
                    // Xử lý click Lưu cho tab Thêm Mới
                    if (window.currentTab === 'themmoi' && window.themMoiQueue && window.themMoiQueue.length > 0) {
                        const btnApproveThemMoi = document.getElementById('btn_approve_themmoi');
                        if (btnApproveThemMoi && btnApproveThemMoi.style.display === 'block' && !btnApproveThemMoi.disabled) {
                            handleApproveThemMoi(true);
                        }
                    }
                }
            }, true); // Sử dụng capture phase (true) để tóm được event trước khi framework Angular nuốt mất sự kiện.
            window.hsskGlobalSaveListenerAdded = true;
        }

        safeAddListener('btn_continue_after_popup', 'click', async () => {
            updateStatus("⏳ Đang dọn dẹp form và sang người mới...");
            setButtonsDisabled(true);
            document.getElementById('btn_continue_after_popup').style.display = 'none';
            const backBtn = findButtonByText('Quay lại');
            if (backBtn) { backBtn.setAttribute('type', 'button'); backBtn.click(); }
            await sleep(300);
            let waitBack = 0;
            while (isWebLoading() && waitBack < 10000) { await sleep(300); waitBack += 300; }
            window.hsskBotQueue.shift();
            updateUI();
            processNextPerson();
        });

        safeAddListener('btn_skip', 'click', async () => {
            if(window.hsskBotQueue.length === 0) return;
            const skipBtn = document.getElementById('btn_skip');
            skipBtn.disabled = true; 
            
            window.currentSearchId = Math.random(); 
            
            updateStatus("⏭️ Đang ngắt tiến trình và bỏ qua...");
            
            const backBtn = findButtonByText('Quay lại');
            if (backBtn) { 
                backBtn.setAttribute('type', 'button'); 
                backBtn.click(); 
            }
            
            await sleep(300);
            let waitBack = 0;
            while (isWebLoading() && waitBack < 5000) { await sleep(300); waitBack += 300; } 
            
            window.hsskBotQueue.shift();
            updateUI();
            skipBtn.disabled = false; 
            processNextPerson();
        });

        safeAddListener('btn_fill_manual', 'click', async () => {
            const btn = document.getElementById('btn_fill_manual');
            const isEditing = Array.from(document.querySelectorAll('input[formcontrolname="currentAddress"]')).some(el => el.offsetParent !== null);
            if (!isEditing) return alert("❌ Bạn chưa mở form Sửa nhân khẩu!");
            btn.innerText = "⏳ Đang điền";
            btn.style.background = "#ff9800";
            btn.disabled = true;
            await dienThongTinVaoForm({});
            btn.innerText = "✅ XONG";
            btn.style.background = "#17a2b8";
            setTimeout(() => { btn.innerText = "✍️ ĐIỀN TAY"; btn.disabled = false; }, 2000);
        });

        safeAddListener('btn_auto_fill_transfer', 'click', async () => {
            const transferDialog = document.querySelector('mat-dialog-container');
            if (!transferDialog || !transferDialog.innerText.includes('Chuyển nhân khẩu về địa bàn')) {
                return alert("Không tìm thấy popup Chuyển địa bàn!");
            }

            updateStatus("✍️ Đang tự động điền form Chuyển địa bàn...");
            const btn = document.getElementById('btn_auto_fill_transfer');
            btn.innerText = "⏳ ĐANG ĐIỀN...";
            btn.disabled = true;

            let person = window.hsskBotQueue.length > 0 ? window.hsskBotQueue[0] : {};
            
            let tinhToFill = person.parsedAddress?.tinh || 'Thành phố Hà Nội';
            let xaToFill = person.parsedAddress?.xa || window.textDienNhanhXa;
            let thonToFill = person.parsedAddress?.thon || window.textDienNhanhThon;
            let thonToSearch = cleanThonXom(thonToFill);
            let diaChiChiTiet = person.parsedAddress?.full || `Thôn ${thonToFill}, ${xaToFill}, ${tinhToFill}`;

            const dialogSelects = transferDialog.querySelectorAll('ng-select');
            if (dialogSelects.length >= 3) {
                await selectNgOption(dialogSelects[0], tinhToFill);
                await selectNgOption(dialogSelects[1], xaToFill);
                if (thonToSearch) await selectNgOption(dialogSelects[2], thonToSearch);
            }

            const addressInput = transferDialog.querySelector('input[formcontrolname="currentAddress"]');
            if (addressInput) {
                addressInput.focus();
                addressInput.value = diaChiChiTiet;
                addressInput.dispatchEvent(new Event('input', { bubbles: true }));
                addressInput.dispatchEvent(new Event('change', { bubbles: true }));
                addressInput.dispatchEvent(new Event('blur', { bubbles: true }));
            }

            const radios = transferDialog.querySelectorAll('label.container-radio');
            for(let r of radios) {
                if(r.innerText.includes('Đã chuẩn hóa')) {
                    const input = r.querySelector('input');
                    if(input && !input.checked) {
                        const checkmark = r.querySelector('.checkmark-radio');
                        if(checkmark) checkmark.click(); else input.click();
                    }
                    break;
                }
            }

            const dateInput = transferDialog.querySelector('input[bsdatepicker], input.cus-datepicker');
            if (dateInput) {
                dateInput.focus();
                dateInput.value = getTodayString();
                dateInput.dispatchEvent(new Event('input', { bubbles: true }));
                dateInput.dispatchEvent(new Event('change', { bubbles: true }));
                dateInput.dispatchEvent(new Event('blur', { bubbles: true }));
            }

            updateStatus("🛑 ĐÃ ĐIỀN XONG CHUYỂN ĐỊA BÀN. HÃY TỰ BẤM [XÁC NHẬN]!");
            btn.innerText = "🏡 TỰ ĐIỀN CHUYỂN ĐỊA BÀN";
            btn.disabled = false;
        });

        safeAddListener('btn_add_cccd_chuanhoa', 'click', async () => {
            const person = window.hsskBotQueue[0];
            if (!person) return alert("Hàng đợi chuẩn hóa đang trống! Vui lòng nạp dữ liệu ở ô trên trước khi dùng chức năng này.");

            let statusEl = document.getElementById('hssk_status');
            setButtonsDisabled(true);

            if (window.location.hash.includes('/sua-nhan-khau')) {
                statusEl.innerHTML = `✍️ Đang ở trang Sửa, tiến hành điền trực tiếp dữ liệu...`;
                
                await dienThongTinVaoForm(person);

                statusEl.innerHTML = `🛑 ĐÃ ĐIỀN XONG. <b style="color:#ff9800; font-size:13px;">HÃY KIỂM TRA VÀ BẤM [DUYỆT & TIẾP TỤC] TRÊN BOT!</b>`;
                document.getElementById('btn_approve').style.display = 'block';
                document.getElementById('btn_skip').style.display = 'block';
                setButtonsDisabled(false);
                setCompactModeUI(false); 
                return; 
            }

            let medId = window.lastCollectedMedicalId;
            if (!medId) {
                medId = prompt("Chưa bắt được Mã y tế tự động. Vui lòng copy và dán Mã y tế vào đây để tiếp tục:");
                if (!medId) { setButtonsDisabled(false); return; }
            }

            statusEl.innerHTML = `🔄 Đang chuyển tab và tìm Mã y tế: ${medId}...`;

            await switchToTab('Nhân khẩu', '/dan-so/ho-gia-dinh/nhan-khau');

            const isSearchReady = fillTextInputByControlName('findBasic', medId);
            if (isSearchReady) {
                await sleep(500);
                const searchBtn2 = findButtonByText('Tìm kiếm');
                if (searchBtn2) {
                    searchBtn2.click();
                    await sleep(500);
                    let waitLoad = 0;
                    while (isWebLoading() && waitLoad < 10000) { await sleep(300); waitLoad += 300; }
                    await sleep(1000);

                    const editRow = document.querySelector('tbody tr');
                    if (editRow && !editRow.innerText.toLowerCase().includes('không')) {
                        const editIcon = editRow.querySelector('mat-icon[svgicon="ic_btn_edit_table"]');
                        if (editIcon) {
                            editIcon.closest('button').click();
                            await sleep(1500);
                            waitLoad = 0;
                            while (isWebLoading() && waitLoad < 10000) { await sleep(300); waitLoad += 300; }

                            statusEl.innerHTML = `✍️ Đang tự động điền thông tin và chuẩn hóa...`;
                            await dienThongTinVaoForm(person);

                            statusEl.innerHTML = `🛑 ĐÃ ĐIỀN XONG. <b style="color:#ff9800; font-size:13px;">HÃY KIỂM TRA VÀ BẤM [DUYỆT & TIẾP TỤC] TRÊN BOT!</b>`;
                            document.getElementById('btn_approve').style.display = 'block';
                            document.getElementById('btn_skip').style.display = 'block';
                            setButtonsDisabled(false);
                            setCompactModeUI(false); 
                            return; 
                        }
                    }
                }
            }

            statusEl.innerHTML = `<span style="color: red;">❌ Không tìm thấy hoặc lỗi. Hãy thử làm thủ công!</span>`;
            setButtonsDisabled(false);
        });

        // ---------------- TAB 2: GHÉP HỘ LISTENERS ----------------
        safeAddListener('btn_import_ghepho', 'click', () => {
            const rawText = document.getElementById('ghepho_excel_input').value.trim();
            if (!rawText) return alert("Vui lòng dán danh sách Excel vào ô trước!");
            
            let lines = rawText.split('\n');
            let map = {};
            
            // Đọc và nhóm theo Tên Chủ Hộ
            lines.forEach(line => {
                if(!line.trim()) return;
                let parts = line.split('\t').map(s => s.trim());
                
                // Format: Tên | Ngày Sinh | Tên Chủ Hộ | CCCD | Địa chỉ | Giới tính
                if(parts.length >= 4) {
                    let name = parts[0];
                    let dob = parts[1];
                    let headName = parts[2] ? parts[2].toUpperCase() : "";
                    let cccd = parts.find(p => /^\d{9,15}$/.test(p) && p.length !== 10) || "";
                    let gender = parts.find(p => /^Nam$|^Nữ$/i.test(p)) || "";
                    
                    if(name && headName && cccd) {
                        if(!map[headName]) map[headName] = [];
                        map[headName].push({name: name, dob: dob, cccd: cccd, gender: gender});
                    }
                }
            });

            window.ghephoQueue = [];
            let errorFamilies = [];
            
            // Xác định ai là chủ, ai là thành viên
            for(let hName in map) {
                let people = map[hName];
                let headIndex = people.findIndex(p => p.name.toUpperCase() === hName);
                
                if(headIndex !== -1) {
                    let head = people[headIndex];
                    people.splice(headIndex, 1); // remove head from list
                    window.ghephoQueue.push({ headName: hName, head: head, members: people });
                } else {
                    errorFamilies.push(hName);
                }
            }
            
            if(errorFamilies.length > 0) {
                alert("Lỗi: Không tìm thấy thông tin của Chủ hộ trong danh sách copy cho các hộ sau:\n" + errorFamilies.join(', ') + "\n\n(Vui lòng kiểm tra lại bôi đen copy có thiếu dòng của chủ hộ không)");
            }
            
            window.totalGhephoFamilies = window.ghephoQueue.length;
            updateUIGhepho();
            document.getElementById('ghepho_excel_input').value = ""; 
            window.textGhepHo = "";
        });

        safeAddListener('btn_start_ghepho', 'click', async () => {
            if (window.ghephoQueue.length === 0) return alert("Hàng đợi hộ gia đình đang trống!");
            document.getElementById('btn_start_ghepho').style.display = 'none';
            document.getElementById('btn_approve_ghepho').style.display = 'block';
            document.getElementById('btn_skip_ghepho').style.display = 'block';
            document.getElementById('btn_stop_ghepho').disabled = false;
            
            window.isGhephoRunning = true;
            processNextGhephoFamily();
        });

        safeAddListener('btn_stop_ghepho', 'click', () => {
            const btn = document.getElementById('btn_stop_ghepho');
            if (window.isGhephoRunning) {
                window.isGhephoRunning = false;
                document.getElementById('ghepho_status').innerHTML = `<span style="color: #ff9800; font-weight: bold;">⏸️ Đã tạm dừng! Bấm Tiếp tục để chạy lại.</span>`;
                btn.innerHTML = '▶️ TIẾP TỤC';
                btn.style.background = '#17a2b8';
            } else {
                window.isGhephoRunning = true;
                document.getElementById('ghepho_status').innerHTML = `<span style="color: green;">▶️ Đang tiếp tục...</span>`;
                btn.innerHTML = '⏸️ TẠM DỪNG';
                btn.style.background = '#dc3545';
                processNextGhephoFamily();
            }
        });

        // --- HÀM XỬ LÝ CHUNG KHI BẤM LƯU GHÉP HỘ (TỪ BOT HOẶC TỪ WEB) ---
        async function handleApproveGhepHo(isNativeClicked = false) {
            let statusEl = document.getElementById('ghepho_status');
            statusEl.innerHTML = isNativeClicked ? "⏳ Phát hiện bấm Lưu trên web. Đang theo dõi..." : "⏳ Đang lưu dữ liệu Hộ Khẩu...";
            setButtonsDisabledGhepho(true);
            
            if (!isNativeClicked) {
                const saveBtn = findButtonByText('Lưu');
                if (saveBtn) {
                    saveBtn.click();
                } else {
                    alert("❌ Không tìm thấy nút Lưu!");
                    setButtonsDisabledGhepho(false);
                    return;
                }
            }

            await sleep(500);
            let waitSave = 0;
            while (isWebLoading() && waitSave < 15000) { await sleep(300); waitSave += 300; }
            await sleep(500); 
            
            // --- BỘ QUÉT TỔNG HỢP (GHÉP HỘ KHÔNG CÓ POPUP TRÙNG) ---
            let isSuccess = false;
            let hasError = false;

            for (let i = 0; i < 30; i++) { 
                if (document.querySelector('.mat-error') || document.querySelector('.invalid-feedback')) {
                    hasError = true; break;
                }
                const toasts = document.querySelectorAll('.toast-message, .cdk-overlay-container, #toast-container, .snack-bar-container, .ngx-toastr');
                for (let toast of toasts) {
                    if (toast && toast.offsetParent !== null) {
                        const text = toast.innerText.toLowerCase();
                        if (text.includes('thành công') || text.includes('cập nhật') || text.includes('lưu')) {
                            isSuccess = true; break;
                        }
                    }
                }
                if(isSuccess) break;
                await sleep(300);
            }

            if (hasError) {
                statusEl.innerHTML = `<span style="color: red;">⚠️ Form báo lỗi (Thiếu trường bắt buộc)! Hãy điền tay rồi bấm Lưu.</span>`;
                setButtonsDisabledGhepho(false); 
                return; 
            } else if (isSuccess) {
                const backBtn = findButtonByText('Quay lại');
                if (backBtn && backBtn.offsetParent !== null) {
                    backBtn.setAttribute('type', 'button'); 
                    backBtn.click();
                    await sleep(300);
                    let waitBack = 0;
                    while (isWebLoading() && waitBack < 10000) { await sleep(300); waitBack += 300; }
                }
                window.ghephoQueue.shift();
                updateUIGhepho();
                processNextGhephoFamily();
            } else {
                statusEl.innerHTML = `<span style="color: red;">⚠️ Không thấy phản hồi từ hệ thống! Bot tạm dừng kiểm tra.</span>`;
                setButtonsDisabledGhepho(false);
                return; 
            }
        }

        safeAddListener('btn_approve_ghepho', 'click', () => {
            handleApproveGhepHo(false);
        });

        safeAddListener('btn_skip_ghepho', 'click', async () => {
            if (window.ghephoQueue.length === 0) return;
            const skipBtn = document.getElementById('btn_skip_ghepho');
            skipBtn.disabled = true; 
            
            window.currentGhephoId = Math.random(); 
            document.getElementById('ghepho_status').innerHTML = "⏭️ Đang bỏ qua hộ này...";
            
            const backBtn = findButtonByText('Quay lại');
            if (backBtn && backBtn.offsetParent !== null) {
                backBtn.setAttribute('type', 'button');
                backBtn.click();
                await sleep(300);
                let waitBack = 0;
                while (isWebLoading() && waitBack < 5000) { await sleep(300); waitBack += 300; }
            }
            
            window.ghephoQueue.shift();
            updateUIGhepho();
            skipBtn.disabled = false; 
            processNextGhephoFamily();
        });


        // ---------------- TAB 3: THÊM MỚI LISTENERS ----------------
        safeAddListener('btn_import_themmoi', 'click', () => {
            const rawText = document.getElementById('auto_themmoi_list').value.trim();
            if (!rawText) return alert("Vui lòng dán danh sách từ Excel vào ô trước!");
            const lines = rawText.split('\n');
            window.themMoiQueue = [];
            lines.forEach(line => {
                let parsed = parseExcelRow(line);
                if (parsed && (parsed.name || parsed.cccd)) window.themMoiQueue.push(parsed);
            });
            window.totalThemMoi = window.themMoiQueue.length;
            updateUIThemMoi();
            document.getElementById('auto_themmoi_list').value = ""; 
            window.textThemMoi = "";
        });

        safeAddListener('btn_start_themmoi', 'click', async () => {
            if (window.themMoiQueue.length === 0) return alert("Danh sách trống!");
            document.getElementById('btn_start_themmoi').style.display = 'none';
            document.getElementById('btn_approve_themmoi').style.display = 'block';
            document.getElementById('btn_skip_themmoi').style.display = 'block';
            document.getElementById('btn_stop_themmoi').disabled = false;
            
            window.isThemMoiRunning = true;
            processNextThemMoi();
        });

        safeAddListener('btn_stop_themmoi', 'click', () => {
            const btn = document.getElementById('btn_stop_themmoi');
            if (window.isThemMoiRunning) {
                window.isThemMoiRunning = false;
                document.getElementById('themmoi_status').innerHTML = `<span style="color: #ff9800; font-weight: bold;">⏸️ Đã tạm dừng! Bấm Tiếp tục để chạy lại.</span>`;
                btn.innerHTML = '▶️ TIẾP TỤC';
                btn.style.background = '#17a2b8';
            } else {
                window.isThemMoiRunning = true;
                document.getElementById('themmoi_status').innerHTML = `<span style="color: green;">▶️ Đang tiếp tục...</span>`;
                btn.innerHTML = '⏸️ TẠM DỪNG';
                btn.style.background = '#dc3545';
                processNextThemMoi();
            }
        });

        // --- HÀM XỬ LÝ CHUNG KHI BẤM LƯU THÊM MỚI (TỪ BOT HOẶC TỪ WEB) ---
        async function handleApproveThemMoi(isNativeClicked = false) {
            let statusEl = document.getElementById('themmoi_status');
            statusEl.innerHTML = isNativeClicked ? "⏳ Phát hiện bấm Lưu trên web. Đang theo dõi..." : "⏳ Đang lưu dữ liệu...";
            setButtonsDisabledThemMoi(true);
            
            if (!isNativeClicked) {
                const saveBtn = findButtonByText('Lưu');
                if (saveBtn) {
                    saveBtn.click();
                } else {
                    alert("❌ Không tìm thấy nút Lưu!");
                    setButtonsDisabledThemMoi(false);
                    return;
                }
            }

            await sleep(500); // Đợi để Loading xuất hiện
            let waitSave = 0;
            while (isWebLoading() && waitSave < 15000) { await sleep(300); waitSave += 300; }
            await sleep(500); 
            
            // --- BỘ QUÉT TỔNG HỢP: TÌM POPUP TRÙNG, TÌM LỖI, VÀ TÌM THÔNG BÁO THÀNH CÔNG CÙNG LÚC ---
            let isSuccess = false;
            let isDuplicate = false;
            let hasError = false;
            let activeDialog = null;

            for (let i = 0; i < 30; i++) { // Quét liên tục trong 9 giây
                // 1. Quét tìm Popup trùng
                const dialog = document.querySelector('mat-dialog-container');
                if (dialog && (dialog.innerText.toLowerCase().includes('nhân khẩu trùng') || dialog.innerText.toLowerCase().includes('xác nhận lưu'))) {
                    isDuplicate = true;
                    activeDialog = dialog;
                    break;
                }

                // 2. Quét tìm Form báo lỗi đỏ
                if (document.querySelector('.mat-error') || document.querySelector('.invalid-feedback')) {
                    hasError = true;
                    break;
                }

                // 3. Quét tìm Toast thông báo thành công
                const toasts = document.querySelectorAll('.toast-message, .cdk-overlay-container, #toast-container, .snack-bar-container, .ngx-toastr');
                for (let toast of toasts) {
                    if (toast && toast.offsetParent !== null) {
                        const text = toast.innerText.toLowerCase();
                        if (text.includes('thành công') || text.includes('cập nhật') || text.includes('lưu') || text.includes('thêm mới')) {
                            isSuccess = true;
                            break;
                        }
                    }
                }
                if(isSuccess) break;

                await sleep(300);
            }

            // --- XỬ LÝ KẾT QUẢ SAU KHI QUÉT ---
            if (isDuplicate) {
                const person = window.themMoiQueue[0];
                const matchResult = await handleDuplicatePopup(person, activeDialog);
                
                if (matchResult && matchResult.handled) {
                    statusEl.innerHTML = `<span style="color: #28a745;">✅ Trùng khớp! Đang chuyển sang trang tra cứu...</span>`;
                    await sleep(1000);
                    let waitLoad = 0;
                    while (isWebLoading() && waitLoad < 10000) { await sleep(300); waitLoad += 300; }
                    
                    await switchToTab('Tra cứu nhân khẩu', '/dan-so/tra-cuu-nhan-khau');

                    let medicalId = ""; 
                    const searchBtn = findButtonByText('Tìm kiếm');
                    if (searchBtn) {
                        statusEl.innerHTML = `🔄 Đang bấm Tìm kiếm...`;
                        searchBtn.click();
                        await sleep(500);
                        waitLoad = 0;
                        while (isWebLoading() && waitLoad < 10000) { await sleep(300); waitLoad += 300; }
                        await sleep(1000);

                        const firstRow = document.querySelector('tbody tr');
                        if (firstRow && !firstRow.innerText.toLowerCase().includes('không')) {
                            const tds = firstRow.querySelectorAll('td');
                            if (tds.length >= 2) {
                                medicalId = tds[1].innerText.trim();
                                window.lastCollectedMedicalId = medicalId; 
                            }

                            const actionBtns = firstRow.querySelectorAll('td:last-child button');
                            if (actionBtns.length > 0) {
                                statusEl.innerHTML = `🔄 Đang mở form Chuyển nhân khẩu...`;
                                actionBtns[0].click(); 
                                await sleep(1500);
                                waitLoad = 0;
                                while (isWebLoading() && waitLoad < 10000) { await sleep(300); waitLoad += 300; }

                                const transferDialog = document.querySelector('mat-dialog-container');
                                if (transferDialog && transferDialog.innerText.includes('Chuyển nhân khẩu về địa bàn')) {
                                    statusEl.innerHTML = `✍️ Đang tự động điền form Chuyển địa bàn...`;
                                    
                                    let tinhToFill = person.parsedAddress?.tinh || 'Thành phố Hà Nội';
                                    let xaToFill = person.parsedAddress?.xa || window.textDienNhanhXa;
                                    let thonToFill = person.parsedAddress?.thon || window.textDienNhanhThon;
                                    let thonToSearch = cleanThonXom(thonToFill);
                                    let diaChiChiTiet = person.parsedAddress?.full || `Thôn ${thonToFill}, ${xaToFill}, ${tinhToFill}`;

                                    const dialogSelects = transferDialog.querySelectorAll('ng-select');
                                    if (dialogSelects.length >= 3) {
                                        await selectNgOption(dialogSelects[0], tinhToFill);
                                        await selectNgOption(dialogSelects[1], xaToFill);
                                        if (thonToSearch) await selectNgOption(dialogSelects[2], thonToSearch);
                                    }

                                    const addressInput = transferDialog.querySelector('input[formcontrolname="currentAddress"]');
                                    if (addressInput) {
                                        addressInput.focus();
                                        addressInput.value = diaChiChiTiet;
                                        addressInput.dispatchEvent(new Event('input', { bubbles: true }));
                                        addressInput.dispatchEvent(new Event('change', { bubbles: true }));
                                        addressInput.dispatchEvent(new Event('blur', { bubbles: true }));
                                    }

                                    const radios = transferDialog.querySelectorAll('label.container-radio');
                                    for(let r of radios) {
                                        if(r.innerText.includes('Đã chuẩn hóa')) {
                                            const input = r.querySelector('input');
                                            if(input && !input.checked) {
                                                const checkmark = r.querySelector('.checkmark-radio');
                                                if(checkmark) checkmark.click(); else input.click();
                                            }
                                            break;
                                        }
                                    }

                                    const dateInput = transferDialog.querySelector('input[bsdatepicker], input.cus-datepicker');
                                    if (dateInput) {
                                        dateInput.focus();
                                        dateInput.value = getTodayString();
                                        dateInput.dispatchEvent(new Event('input', { bubbles: true }));
                                        dateInput.dispatchEvent(new Event('change', { bubbles: true }));
                                        dateInput.dispatchEvent(new Event('blur', { bubbles: true }));
                                    }

                                    statusEl.innerHTML = `🛑 ĐÃ ĐIỀN XONG CHUYỂN ĐỊA BÀN. <b style="color: #ff9800; font-size: 13px;">HÃY TỰ BẤM [XÁC NHẬN] TRÊN MÀN HÌNH!</b>`;
                                    
                                    while (document.querySelector('mat-dialog-container')) {
                                        if (!window.isThemMoiRunning) return; 
                                        await sleep(500);
                                    }

                                    waitLoad = 0;
                                    while (isWebLoading() && waitLoad < 10000) { await sleep(300); waitLoad += 300; }
                                }
                            }
                        }

                        if (!matchResult.isCccdMatched && medicalId) {
                             statusEl.innerHTML = `<span style="color: #ff9800;">⚠️ Hệ thống đang thiếu CCCD. Bạn có thể bấm nút Tím bên dưới để Bot tự sửa bổ sung!</span>`;
                             setButtonsDisabledThemMoi(false);
                             
                             const btnStop = document.getElementById('btn_stop_themmoi');
                             window.isThemMoiRunning = false;
                             btnStop.innerHTML = '▶️ TIẾP TỤC (BỎ QUA SỬA)';
                             btnStop.style.background = '#17a2b8';
                             setCompactModeUI(false);
                             return; 
                        }

                        await switchToTab('Thêm mới nhân khẩu', '/dan-so/ho-gia-dinh/nhan-khau/nhan-khau-them-moi');

                        statusEl.innerHTML = `<span style="color: #28a745;">✅ Xong tác vụ xử lý người trùng!</span>`;
                        setCompactModeUI(true);
                        window.themMoiQueue.shift();
                        updateUIThemMoi();
                        processNextThemMoi(); 
                        return;
                    } // <-- FIX LỖI: CẶP NGOẶC BỊ THIẾU Ở ĐÂY ĐÃ ĐƯỢC THÊM VÀO
                } else {
                    statusEl.innerHTML = `<span style="color: red;">⚠️ ${matchResult.reason}. Hãy tự quyết định!</span>`;
                    setButtonsDisabledThemMoi(false); 
                    
                    const btnStop = document.getElementById('btn_stop_themmoi');
                    window.isThemMoiRunning = false;
                    btnStop.innerHTML = '▶️ TIẾP TỤC (BỎ QUA SỬA)';
                    btnStop.style.background = '#17a2b8';
                    setCompactModeUI(false);
                    return; 
                }
            } else if (hasError) {
                statusEl.innerHTML = `<span style="color: red;">⚠️ Web báo lỗi (Thiếu thông tin)! Hãy xử lý tay rồi bấm Lưu.</span>`;
                setButtonsDisabledThemMoi(false); 
                setCompactModeUI(false);
                return; 
            } else if (isSuccess) {
                setCompactModeUI(true);
                const backBtn = findButtonByText('Quay lại');
                if (backBtn && backBtn.offsetParent !== null) {
                    backBtn.setAttribute('type', 'button'); 
                    backBtn.click();
                    await sleep(300);
                    let waitBack = 0;
                    while (isWebLoading() && waitBack < 10000) { await sleep(300); waitBack += 300; }
                }

                window.themMoiQueue.shift();
                updateUIThemMoi();
                processNextThemMoi();

            } else {
                statusEl.innerHTML = `<span style="color: red;">⚠️ Không thấy phản hồi từ hệ thống! Bot tạm dừng để bạn kiểm tra.</span>`;
                setButtonsDisabledThemMoi(false);
                setCompactModeUI(false);
            }
        }

        safeAddListener('btn_approve_themmoi', 'click', () => {
            handleApproveThemMoi(false);
        });

        safeAddListener('btn_skip_themmoi', 'click', async () => {
            if (window.themMoiQueue.length === 0) return;
            const skipBtn = document.getElementById('btn_skip_themmoi');
            skipBtn.disabled = true; 
            
            window.currentThemMoiId = Math.random(); 
            
            document.getElementById('themmoi_status').innerHTML = "⏭️ Đang ngắt tiến trình và bỏ qua...";
            
            const backBtn = findButtonByText('Quay lại');
            if (backBtn && backBtn.offsetParent !== null) {
                backBtn.setAttribute('type', 'button');
                backBtn.click();
                await sleep(300);
                let waitBack = 0;
                while (isWebLoading() && waitBack < 5000) { await sleep(300); waitBack += 300; }
            }
            
            window.themMoiQueue.shift();
            updateUIThemMoi();
            skipBtn.disabled = false; 
            processNextThemMoi();
        });

        safeAddListener('btn_add_cccd_themmoi', 'click', async () => {
            const person = window.themMoiQueue[0];
            if (!person) return alert("Hàng đợi thêm mới đang trống! Vui lòng nạp dữ liệu ở ô trên trước khi dùng chức năng này.");

            let statusEl = document.getElementById('themmoi_status');
            setButtonsDisabledThemMoi(true);

            if (window.location.hash.includes('/sua-nhan-khau')) {
                statusEl.innerHTML = `✍️ Đang ở trang Sửa, tiến hành điền trực tiếp dữ liệu...`;
                
                await dienThongTinVaoForm(person);

                statusEl.innerHTML = `🛑 ĐÃ ĐIỀN XONG. <b style="color:#ff9800; font-size:13px;">HÃY KIỂM TRA VÀ BẤM [DUYỆT & LƯU] TRÊN BOT!</b>`;
                document.getElementById('btn_approve_themmoi').style.display = 'block';
                document.getElementById('btn_skip_themmoi').style.display = 'block';
                setButtonsDisabledThemMoi(false);
                setCompactModeUI(false); 
                return; 
            }

            let medId = window.lastCollectedMedicalId;
            if (!medId) {
                medId = prompt("Chưa bắt được Mã y tế tự động. Vui lòng copy và dán Mã y tế vào đây để tiếp tục:");
                if (!medId) { setButtonsDisabledThemMoi(false); return; }
            }

            statusEl.innerHTML = `🔄 Đang chuyển tab và tìm Mã y tế: ${medId}...`;

            await switchToTab('Nhân khẩu', '/dan-so/ho-gia-dinh/nhan-khau');

            const isSearchReady = fillTextInputByControlName('findBasic', medId);
            if (isSearchReady) {
                await sleep(500);
                const searchBtn2 = findButtonByText('Tìm kiếm');
                if (searchBtn2) {
                    searchBtn2.click();
                    await sleep(500);
                    let waitLoad = 0;
                    while (isWebLoading() && waitLoad < 10000) { await sleep(300); waitLoad += 300; }
                    await sleep(1000);

                    const editRow = document.querySelector('tbody tr');
                    if (editRow && !editRow.innerText.toLowerCase().includes('không')) {
                        const editIcon = editRow.querySelector('mat-icon[svgicon="ic_btn_edit_table"]');
                        if (editIcon) {
                            editIcon.closest('button').click();
                            await sleep(1500);
                            waitLoad = 0;
                            while (isWebLoading() && waitLoad < 10000) { await sleep(300); waitLoad += 300; }

                            statusEl.innerHTML = `✍️ Đang tự động điền thông tin và chuẩn hóa...`;
                            await dienThongTinVaoForm(person);

                            statusEl.innerHTML = `🛑 ĐÃ ĐIỀN XONG. <b style="color:#ff9800; font-size:13px;">HÃY KIỂM TRA VÀ BẤM [DUYỆT & LƯU] TRÊN BOT!</b>`;
                            document.getElementById('btn_approve_themmoi').style.display = 'block';
                            document.getElementById('btn_skip_themmoi').style.display = 'block';
                            setButtonsDisabledThemMoi(false);
                            setCompactModeUI(false); 
                            return; 
                        }
                    }
                }
            }

            statusEl.innerHTML = `<span style="color: red;">❌ Không tìm thấy hoặc lỗi. Hãy thử làm thủ công!</span>`;
            setButtonsDisabledThemMoi(false);
        });

        safeAddListener('btn_auto_fill_transfer_themmoi', 'click', async () => {
            const transferDialog = document.querySelector('mat-dialog-container');
            if (!transferDialog || !transferDialog.innerText.includes('Chuyển nhân khẩu về địa bàn')) {
                return alert("Không tìm thấy popup Chuyển địa bàn!");
            }

            let statusEl = document.getElementById('themmoi_status');
            if(statusEl) statusEl.innerHTML = "✍️ Đang tự động điền form Chuyển địa bàn...";
            
            const btn = document.getElementById('btn_auto_fill_transfer_themmoi');
            btn.innerText = "⏳ ĐANG ĐIỀN...";
            btn.disabled = true;

            let person = window.themMoiQueue.length > 0 ? window.themMoiQueue[0] : {};
            
            let tinhToFill = person.parsedAddress?.tinh || 'Thành phố Hà Nội';
            let xaToFill = person.parsedAddress?.xa || window.textDienNhanhXa;
            let thonToFill = person.parsedAddress?.thon || window.textDienNhanhThon;
            let thonToSearch = cleanThonXom(thonToFill);
            let diaChiChiTiet = person.parsedAddress?.full || `Thôn ${thonToFill}, ${xaToFill}, ${tinhToFill}`;

            const dialogSelects = transferDialog.querySelectorAll('ng-select');
            if (dialogSelects.length >= 3) {
                await selectNgOption(dialogSelects[0], tinhToFill);
                await selectNgOption(dialogSelects[1], xaToFill);
                if (thonToSearch) await selectNgOption(dialogSelects[2], thonToSearch);
            }

            const addressInput = transferDialog.querySelector('input[formcontrolname="currentAddress"]');
            if (addressInput) {
                addressInput.focus();
                addressInput.value = diaChiChiTiet;
                addressInput.dispatchEvent(new Event('input', { bubbles: true }));
                addressInput.dispatchEvent(new Event('change', { bubbles: true }));
                addressInput.dispatchEvent(new Event('blur', { bubbles: true }));
            }

            const radios = transferDialog.querySelectorAll('label.container-radio');
            for(let r of radios) {
                if(r.innerText.includes('Đã chuẩn hóa')) {
                    const input = r.querySelector('input');
                    if(input && !input.checked) {
                        const checkmark = r.querySelector('.checkmark-radio');
                        if(checkmark) checkmark.click(); else input.click();
                    }
                    break;
                }
            }

            const dateInput = transferDialog.querySelector('input[bsdatepicker], input.cus-datepicker');
            if (dateInput) {
                dateInput.focus();
                dateInput.value = getTodayString();
                dateInput.dispatchEvent(new Event('input', { bubbles: true }));
                dateInput.dispatchEvent(new Event('change', { bubbles: true }));
                dateInput.dispatchEvent(new Event('blur', { bubbles: true }));
            }

            if(statusEl) statusEl.innerHTML = "🛑 ĐÃ ĐIỀN XONG CHUYỂN ĐỊA BÀN. HÃY TỰ BẤM [XÁC NHẬN]!";
            btn.innerText = "🏡 TỰ ĐIỀN CHUYỂN ĐỊA BÀN";
            btn.disabled = false;
        });

    }

    // ==========================================
    // CÁC HÀM TIỆN ÍCH CHUNG & QUÉT MÃ Y TẾ NGẦM
    // ==========================================
    window.isSuggestListenerAdded = window.isSuggestListenerAdded || false;

    function monitorMedicalId() {
        let btnChuanHoa = document.getElementById('btn_add_cccd_chuanhoa');
        let btnThemMoi = document.getElementById('btn_add_cccd_themmoi');
        if (!btnChuanHoa && !btnThemMoi) return;

        if (window.location.hash.includes('/sua-nhan-khau')) {
            const text = "🏷️ ĐIỀN TRỰC TIẾP VÀO FORM SỬA";
            if (btnChuanHoa && btnChuanHoa.innerText !== text) {
                btnChuanHoa.innerText = text;
                btnChuanHoa.style.background = '#28a745'; 
                btnChuanHoa.title = "Bấm để điền trực tiếp thông tin vào form này";
            }
            if (btnThemMoi && btnThemMoi.innerText !== text) {
                btnThemMoi.innerText = text;
                btnThemMoi.style.background = '#28a745';
                btnThemMoi.title = "Bấm để điền trực tiếp thông tin vào form này";
            }
            return; 
        }

        let foundId = null;

        const transferDialog = document.querySelector('mat-dialog-container');
        if (transferDialog && transferDialog.innerText.includes('Chuyển nhân khẩu về địa bàn')) {
            const text = transferDialog.innerText;
            const match = text.match(/\b(0\d{9,15})\b/);
            if (match) {
                foundId = match[1];
            }
        }

        if (!foundId && window.location.hash.includes('/dan-so/tra-cuu-nhan-khau')) {
            const rows = document.querySelectorAll('tbody tr');
            if (rows.length > 0 && !rows[0].innerText.toLowerCase().includes('không')) {
                const tds = rows[0].querySelectorAll('td');
                if (tds.length >= 2) {
                    let potentialId = tds[1].innerText.trim();
                    if (potentialId && /\d/.test(potentialId) && potentialId.length >= 5) {
                        foundId = potentialId;
                    }
                }
            }
        }

        if (foundId) {
            window.lastCollectedMedicalId = foundId;
            let newText = `🏷️ SỬA MÃ: ${foundId}`;
            
            if (btnChuanHoa && btnChuanHoa.innerText !== newText) {
                btnChuanHoa.innerText = newText;
                btnChuanHoa.style.background = '#d81b60'; 
                btnChuanHoa.title = `Mã y tế đang lưu: ${foundId}`;
            }
            if (btnThemMoi && btnThemMoi.innerText !== newText) {
                btnThemMoi.innerText = newText;
                btnThemMoi.style.background = '#d81b60';
                btnThemMoi.title = `Mã y tế đang lưu: ${foundId}`;
            }
        } else if (window.lastCollectedMedicalId) {
            let oldText = `🏷️ SỬA MÃ: ${window.lastCollectedMedicalId}`;
            if (btnChuanHoa && btnChuanHoa.innerText !== oldText) {
                btnChuanHoa.innerText = oldText;
                btnChuanHoa.style.background = '#d81b60';
                btnChuanHoa.title = `Mã y tế đang lưu: ${window.lastCollectedMedicalId}`;
            }
            if (btnThemMoi && btnThemMoi.innerText !== oldText) {
                btnThemMoi.innerText = oldText;
                btnThemMoi.style.background = '#d81b60';
                btnThemMoi.title = `Mã y tế đang lưu: ${window.lastCollectedMedicalId}`;
            }
        } else {
            let defaultText = "🏷️ SỬA BẰNG MÃ Y TẾ";
            if (btnChuanHoa && btnChuanHoa.innerText !== defaultText) {
                btnChuanHoa.innerText = defaultText;
                btnChuanHoa.style.background = '#e83e8c';
                btnChuanHoa.title = "Chưa có mã y tế nào được lưu";
            }
            if (btnThemMoi && btnThemMoi.innerText !== defaultText) {
                btnThemMoi.innerText = defaultText;
                btnThemMoi.style.background = '#e83e8c';
                btnThemMoi.title = "Chưa có mã y tế nào được lưu";
            }
        }
    }

    function initAutoSuggest() {
        let rawData1 = window.textChuanHoa || '';
        let rawData2 = window.textGhepHo || '';
        let rawData3 = window.textThemMoi || '';

        let combinedText = rawData1 + '\n' + rawData2 + '\n' + rawData3;
        let lines = combinedText.split('\n');
        let suggestions = new Map();

        lines.forEach(line => {
            line = line.trim();
            if (!line) return;
            let cccd = "";
            let parts = line.split(/[\t|,|;]/); 
            if(parts.length > 1) {
               let found = parts.find(p => /^[a-zA-Z0-9]{9,15}$/.test(p.trim()) && /\d/.test(p.trim()));
               if(found) cccd = found.trim();
            } else {
               let match = line.match(/\b([a-zA-Z0-9]{9,15})\b/);
               if(match && /\d/.test(match[1])) cccd = match[1];
            }
            if (cccd) {
                let desc = line.replace(/\t/g, ' - ').substring(0, 80); 
                suggestions.set(cccd, desc);
            }
        });

        let dataList = document.getElementById('hssk_cccd_suggestions');
        if (!dataList) {
            dataList = document.createElement('datalist');
            dataList.id = 'hssk_cccd_suggestions';
            document.body.appendChild(dataList);
        }

        let newHTML = '';
        suggestions.forEach((desc, cccd) => { newHTML += `<option value="${cccd}">${desc}</option>`; });
        if (dataList.innerHTML !== newHTML) dataList.innerHTML = newHTML;

        if (!window.isSuggestListenerAdded) {
            document.addEventListener('input', function(e) {
                if (e.target && e.target.tagName === 'INPUT' && (e.target.type === 'text' || e.target.type === 'search' || !e.target.type)) {
                    const val = e.target.value;
                    if (/^0\d*$/.test(val)) {
                        if (e.target.getAttribute('list') !== 'hssk_cccd_suggestions') {
                            e.target.setAttribute('list', 'hssk_cccd_suggestions');
                            e.target.autocomplete = 'off';
                        }
                    } else {
                        if (e.target.getAttribute('list') === 'hssk_cccd_suggestions') e.target.removeAttribute('list');
                    }
                }
            });
            window.isSuggestListenerAdded = true;
        }
    }

    function applyCompactMode() {
        const labelsToHide = [
            'Quốc tịch', 'Dân tộc', 'Tôn giáo', 'Nghề nghiệp', 'Nơi cấp', 'Hộ chiếu', 'Điện thoại cơ quan', 'Email', 'Ngày chuyển đến',
            'Trình độ học vấn', 'Nơi làm việc', 'Tình trạng hôn nhân', 'Mã số thẻ người khuyết tật', 'Nơi công tác/Học tập', 'Nơi làm việc khác', 'Người chăm sóc',
            'Mã định danh y tế bố', 'Họ tên bố', 'Số CCCD bố', 'Số điện thoại bố', 'Mã định danh y tế mẹ', 'Họ tên mẹ', 'Số CCCD mẹ', 'Số điện thoại mẹ',
            'Họ tên người C/S chính', 'Mối quan hệ với chủ hộ', 'Số CCCD người C/S chính', 'Số điện thoại liên hệ', 'Mã hộ khẩu', 'Họ tên chủ hộ'
        ];
        const formGroups = document.querySelectorAll('.form-group');
        formGroups.forEach(group => {
            const label = group.querySelector('label');
            if (label) {
                let labelText = label.innerText.trim().replace(/\*/g, '').trim(); 
                let shouldHide = labelsToHide.some(hideText => labelText.includes(hideText));
                if (labelText === 'Số CMND' || labelText === 'Ngày cấp') shouldHide = true;
                if (shouldHide) group.style.display = window.isCompactMode ? 'none' : '';
            }
        });
    }

    function setCompactModeUI(isCompact) {
        window.isCompactMode = isCompact;
        const toggle = document.getElementById('toggle_compact');
        if(toggle) toggle.checked = isCompact;
        applyCompactMode();
    }

    window.hsskBotInterval = setInterval(() => {
        if (!document.querySelector('.app-loading')) {
            createBotPanel();
            initAutoSuggest();
            applyCompactMode(); 
            monitorMedicalId(); 

            const isTraCuuUrl = window.location.hash.includes('/dan-so/tra-cuu-nhan-khau');
            const transferDialog = document.querySelector('mat-dialog-container');
            const isTransferPopupOpen = transferDialog && transferDialog.innerText.includes('Chuyển nhân khẩu về địa bàn');
            
            const btnTransferChuanHoa = document.getElementById('btn_auto_fill_transfer');
            if (btnTransferChuanHoa) {
                if (isTraCuuUrl && isTransferPopupOpen && window.currentTab === 'chuanhoa') {
                    btnTransferChuanHoa.style.display = 'block';
                } else {
                    btnTransferChuanHoa.style.display = 'none';
                }
            }

            const btnTransferThemMoi = document.getElementById('btn_auto_fill_transfer_themmoi');
            if (btnTransferThemMoi) {
                if (isTraCuuUrl && isTransferPopupOpen && window.currentTab === 'themmoi') {
                    btnTransferThemMoi.style.display = 'block';
                } else {
                    btnTransferThemMoi.style.display = 'none';
                }
            }
        }
    }, 2000);

    const sleep = (ms) => new Promise(resolve => setTimeout(resolve, ms));

    async function processNextPerson() {
        if (window.hsskBotQueue.length === 0) {
            updateStatus("🎉 ĐÃ HOÀN THÀNH TOÀN BỘ DANH SÁCH!");
            const btnStart = document.getElementById('btn_auto_start');
            if(btnStart) btnStart.style.display = 'block';
            const btnApp = document.getElementById('btn_approve');
            if(btnApp) btnApp.style.display = 'none';
            const btnSkip = document.getElementById('btn_skip');
            if(btnSkip) btnSkip.style.display = 'none';
            const btnCont = document.getElementById('btn_continue_after_popup');
            if(btnCont) btnCont.style.display = 'none';
            return;
        }
        
        const searchId = Math.random();
        window.currentSearchId = searchId;
        
        const person = window.hsskBotQueue[0];
        updateUI();

        const btnApp = document.getElementById('btn_approve');
        if(btnApp) btnApp.style.display = 'none';
        const btnSkip = document.getElementById('btn_skip');
        if(btnSkip) btnSkip.style.display = 'block'; 
        setButtonsDisabled(true);

        if (!window.location.hash.includes('/dan-so/ho-gia-dinh/nhan-khau')) {
            updateStatus("🔄 Đang chuyển về màn hình tra cứu...");
            window.location.hash = '/dan-so/ho-gia-dinh/nhan-khau';
            await sleep(500); 
            let waitLoading = 0;
            while (isWebLoading()) { await sleep(300); waitLoading += 300; if (waitLoading > 15000) break; }
        }

        const isEditing = Array.from(document.querySelectorAll('input[formcontrolname="currentAddress"]')).some(el => el.offsetParent !== null);
        if (isEditing) {
            const backBtn = findButtonByText('Quay lại');
            if (backBtn) { backBtn.setAttribute('type', 'button'); backBtn.click(); }
            await sleep(300);
            let waitLoading = 0;
            while (isWebLoading()) { await sleep(300); waitLoading += 300; if (waitLoading > 10000) break; }
        }

        if (window.currentSearchId !== searchId) return; 

        updateStatus(`🔍 Đang tra CCCD: ${person.cccd}`);
        let found = await searchAndOpenCorrectPerson(person);

        if (window.currentSearchId !== searchId) return; 

        if (found) {
            updateUI(); 
            updateStatus(`✍️ Đang tự động điền Form...`);
            setButtonsDisabled(true);
            await dienThongTinVaoForm(person);
            
            if (window.currentSearchId !== searchId) return; 
            
            updateUI(); 
            updateStatus(`🛑 ĐÃ ĐIỀN XONG. HÃY KIỂM TRA VÀ BẤM DUYỆT!`);
            const bApp = document.getElementById('btn_approve');
            if(bApp) bApp.style.display = 'block';
            setButtonsDisabled(false); 
        } else {
            updateStatus(`❌ Không tìm thấy ${person.cccd}. Đang chuyển sang tab THÊM MỚI...`);
            await sleep(1500); 
            if (window.currentSearchId === searchId) {
                const notFoundPerson = window.hsskBotQueue.shift();
                window.themMoiQueue.push(notFoundPerson);
                window.totalThemMoi++; 
                updateUIThemMoi();
                updateUI();
                processNextPerson();
            }
        }
    }

    async function searchAndOpenCorrectPerson(person) {
        const mySearchId = window.currentSearchId; 
        const isSearchInputFound = fillTextInputByControlName('findBasic', person.cccd);
        if (!isSearchInputFound) return false;
        await sleep(300); 
        const searchBtn = findButtonByText('Tìm kiếm');
        if (!searchBtn) return false;
        searchBtn.click();
        await sleep(500); 
        let waitLoading = 0;
        while (isWebLoading()) {
            if (window.currentSearchId !== mySearchId) return false;
            updateStatus(`⏳ Máy chủ đang tải dữ liệu (${waitLoading/1000}s)...`);
            await sleep(1000); waitLoading += 1000; if (waitLoading > 90000) break;
        }

        let isListReady = false;
        for (let i = 0; i < 40; i++) { 
            if (window.currentSearchId !== mySearchId) return false;
            await sleep(500);
            if (isWebLoading()) continue;
            const rows = Array.from(document.querySelectorAll('tbody tr'));
            if (rows.length === 1 && rows[0].innerText.toLowerCase().includes('không')) return false; 
            const totalText = Array.from(document.querySelectorAll('span.total-table')).map(s => s.innerText.trim()).join(' ');
            const match = totalText.match(/\(([\d,.]+)\)/);
            if (match) {
                const totalNum = parseInt(match[1].replace(/[,.]/g, ''), 10);
                if (totalNum > 0) { isListReady = true; break; }
            }
        }

        if (!isListReady) return false;
        await sleep(600); 

        const rows = Array.from(document.querySelectorAll('tbody tr')).filter(r => r.offsetParent !== null);
        if (rows.length > 0 && !rows[0].innerText.toLowerCase().includes('không')) {
            const targetBtn = rows[0].querySelector('mat-icon[svgicon="ic_btn_edit_table"]');
            if (targetBtn) {
                let rowTxt = rows[0].innerText.toLowerCase();
                if (rowTxt.match(/\bnam\b/)) person.gioiTinhTraCuu = 'Nam';
                else if (rowTxt.match(/\bnữ\b/)) person.gioiTinhTraCuu = 'Nữ';
                let matchDob = rowTxt.match(/\d{2}\/\d{2}\/(\d{4})/);
                if (matchDob) person.ngaySinhTraCuu = matchDob[0];
                let aTag = rows[0].querySelector('td a');
                if(aTag) {
                    let td = aTag.closest('td');
                    if (td && td.nextElementSibling) person.hoTenTraCuu = td.nextElementSibling.innerText.trim();
                }
                targetBtn.scrollIntoView({ behavior: 'smooth', block: 'center' });
                await sleep(300);
                targetBtn.closest('button').click();
                updateStatus(`⏳ Đang mở Form điền thông tin...`);
                for (let i = 0; i < 30; i++) {
                    if (window.currentSearchId !== mySearchId) return false;
                    await sleep(400);
                    const isReady = Array.from(document.querySelectorAll('input[formcontrolname="currentAddress"]')).some(el => el.offsetParent !== null);
                    if (isReady) { await sleep(500); return true; }
                }
            }
        }
        return false;
    }


    // ===============================================
    // HÀM XỬ LÝ GHÉP HỘ MỚI
    // ===============================================

    function updateUIGhepho() {
        const queueCountEl = document.getElementById('ghepho_queue_count');
        const queueTotalEl = document.getElementById('ghepho_total_count');
        if (queueCountEl) queueCountEl.innerText = window.ghephoQueue.length;
        if (queueTotalEl) queueTotalEl.innerText = window.totalGhephoFamilies;
        
        const currentFamilyEl = document.getElementById('ghepho_current_family');
        if (currentFamilyEl) {
            if (window.ghephoQueue.length > 0) {
                const f = window.ghephoQueue[0];
                let htmlStr = `🏠 Chủ hộ: <b style="color:#ff9800">${f.headName}</b> (${f.head.cccd})<br>`;
                htmlStr += `👥 Số thành viên phụ thuộc: <b>${f.members.length}</b> người`;
                currentFamilyEl.innerHTML = htmlStr;
            } else {
                currentFamilyEl.innerHTML = "Trống";
            }
        }
    }

    function setButtonsDisabledGhepho(isDisabled) {
        const btnApprove = document.getElementById('btn_approve_ghepho');
        if(btnApprove) { btnApprove.disabled = isDisabled; btnApprove.style.opacity = isDisabled ? "0.5" : "1"; }
    }

    async function processNextGhephoFamily() {
        if (!window.isGhephoRunning) return;
        
        if (window.ghephoQueue.length === 0) {
            document.getElementById('ghepho_status').innerHTML = `<span style="color: green; font-weight: bold;">🎉 ĐÃ GHÉP XONG TOÀN BỘ DANH SÁCH!</span>`;
            document.getElementById('btn_start_ghepho').style.display = 'block';
            document.getElementById('btn_approve_ghepho').style.display = 'none';
            document.getElementById('btn_skip_ghepho').style.display = 'none';
            document.getElementById('btn_stop_ghepho').disabled = true;
            window.isGhephoRunning = false;
            return;
        }

        const family = window.ghephoQueue[0];
        updateUIGhepho();
        let statusEl = document.getElementById('ghepho_status');
        setButtonsDisabledGhepho(true);

        const runId = Math.random();
        window.currentGhephoId = runId;

        // Chuyển tới form Thêm mới Hộ Khẩu
        if (!window.location.hash.includes('/dan-so/ho-gia-dinh/ho-khau/them-moi-ho-khau')) {
            statusEl.innerHTML = "🔄 Đang mở trang Thêm mới Hộ khẩu...";
            window.location.hash = '/dan-so/ho-gia-dinh/ho-khau/them-moi-ho-khau';
            await sleep(500); 
            let waitLoading = 0;
            while (isWebLoading()) { await sleep(300); waitLoading += 300; if (waitLoading > 15000) break; } 
            await sleep(500); 
        }

        if (window.currentGhephoId !== runId || !window.isGhephoRunning) return;

        // Bắt đầu nhồi người vào UI
        let resultLog = await executeGhephoForFamily(family, statusEl);

        if (window.currentGhephoId !== runId || !window.isGhephoRunning) return;

        statusEl.innerHTML = `🛑 ĐÃ GHÉP XONG HỘ <b>${family.headName}</b>.<br><div style="background:#fff3cd; padding:6px; border-radius:4px; margin:6px 0; font-size:12px; line-height:1.4;">${resultLog}</div>HÃY KIỂM TRA VÀ BẤM DUYỆT!`;
        document.getElementById('btn_approve_ghepho').style.display = 'block';
        document.getElementById('btn_skip_ghepho').style.display = 'block';
        setButtonsDisabledGhepho(false);
    }

    async function executeGhephoForFamily(family, statusEl) {
        // Tạo chuỗi task: Đầu tiên là Chủ hộ, sau đó là Thành viên
        let tasks = [];
        tasks.push({ role: 'chu_ho', cccd: family.head.cccd, name: family.head.name });
        family.members.forEach(m => {
            tasks.push({ role: 'thanh_vien', cccd: m.cccd, name: m.name });
        });

        const closePopupIfOpen = async () => {
            let btnClose = document.querySelector('mat-icon[svgicon="ic_closePeople"]')?.closest('button') || document.querySelector('.button-close');
            if (btnClose) { btnClose.click(); await sleep(500); }
        };

        for (let i = 0; i < tasks.length; i++) {
            if (!window.isGhephoRunning) break;
            let currentTask = tasks[i];
            let cccd = currentTask.cccd;
            let isAddingChuHo = (currentTask.role === 'chu_ho');

            statusEl.innerHTML = `Đang tìm và thêm: <b>${currentTask.name}</b> (${cccd})...`;

            try {
                let popupContainer = document.querySelector('app-search-list-patients');
                if (!popupContainer) {
                    if (isAddingChuHo) {
                        const btnChonChuHo = Array.from(document.querySelectorAll('button.btn-choose')).find(b => b.textContent.trim() === 'Chọn');
                        if (btnChonChuHo) btnChonChuHo.click();
                    } else {
                        let btnAddPeople = findButtonByText('Thêm nhân khẩu');
                        if (btnAddPeople) btnAddPeople.click();
                    }
                    await sleep(1000);
                    popupContainer = document.querySelector('app-search-list-patients');
                }

                if (!popupContainer) { continue; }

                let searchInput = popupContainer.querySelector('input[formcontrolname="findBasic"]');
                if (searchInput) {
                    searchInput.focus(); searchInput.value = cccd; searchInput.dispatchEvent(new Event('input', { bubbles: true }));
                    await sleep(300); 
                }

                let btnSearchSpan = Array.from(popupContainer.querySelectorAll('span.text, button')).find(s => s.innerText.trim() === 'Tìm kiếm');
                if (btnSearchSpan) {
                    let btnSearch = btnSearchSpan.tagName === 'BUTTON' ? btnSearchSpan : btnSearchSpan.closest('button');
                    if (btnSearch) {
                        btnSearch.click(); await sleep(400);
                        let waitLoading = 0;
                        while (isWebLoading()) { if (!window.isGhephoRunning) return; await sleep(500); waitLoading += 500; if (waitLoading > 30000) break; }
                        await sleep(1000); 
                    }
                }

                if (isAddingChuHo) {
                    let btnIconChon = popupContainer.querySelector('table tbody tr td button.btn-choose');
                    if (btnIconChon) {
                        btnIconChon.click(); await sleep(700); continue; 
                    } else {
                        await closePopupIfOpen(); continue;
                    }
                } else {
                    let checkbox = popupContainer.querySelector('table tbody tr input[type="checkbox"]');
                    if (checkbox) {
                        if (!checkbox.checked) checkbox.click();
                        await sleep(300);
                        let btnConfirmSpan = Array.from(popupContainer.querySelectorAll('span.text, button')).find(s => s.innerText.trim() === 'Xác nhận');
                        if (btnConfirmSpan) {
                            let btnConfirm = btnConfirmSpan.tagName === 'BUTTON' ? btnConfirmSpan : btnConfirmSpan.closest('button');
                            if (btnConfirm) { btnConfirm.click(); await sleep(1000); }
                        }
                        await closePopupIfOpen();
                    } else {
                        await closePopupIfOpen();
                    }
                }
            } catch (error) { console.error(error); await closePopupIfOpen(); }
        }

        // TỰ ĐỘNG CHỌN QUAN HỆ GIA ĐÌNH DỰA VÀO ĐỘ TUỔI VÀ GIỚI TÍNH
        statusEl.innerHTML = `⚙️ Đang dùng Logic để tính toán Quan hệ gia đình...`;
        
        // Hàm hỗ trợ lấy Họ
        const layHo = (hoTen) => hoTen ? hoTen.trim().split(' ')[0].toUpperCase() : "";
        
        // 1. Gán vai trò sơ bộ (đoán theo tuổi & giới tính) cho tất cả thành viên
        family.members.forEach(m => {
            m.vaiTro = guessRelationship(family.head, m);
        });

        // 2. Xác định "Họ chuẩn" của gia đình
        let hoChuan = "";
        let hGender = family.head.gender ? family.head.gender.toLowerCase() : "";
        if (hGender === 'nam') {
            hoChuan = layHo(family.head.name);
        } else {
            let chong = family.members.find(m => m.vaiTro === 'Chồng');
            if (chong) hoChuan = layHo(chong.name);
        }

        // 3. Chạy tinh chỉnh bắt lỗi khác Họ
        family.members.forEach(m => {
            if ((m.vaiTro === 'Con trai' || m.vaiTro === 'Con gái') && hoChuan !== "") {
                let hoThanhVien = layHo(m.name);
                if (hoThanhVien !== hoChuan) {
                    let mGen = m.gender ? m.gender.toLowerCase() : "";
                    if (mGen === 'nữ') m.vaiTro = 'Con dâu';
                    else if (mGen === 'nam') m.vaiTro = 'Con rể';
                }
            }
        });
        // ----------------------------------------------------------------------

        let rows = document.querySelectorAll('tbody tr');
        let logText = `<b>⚙️ Logic dự đoán:</b><br>• ${family.headName} (${family.head.dob || 'Không rõ NS'}): <b style="color:#d81b60">Chủ hộ</b><br>`;
        for (let row of rows) {
            if (row.offsetParent === null) continue;
            
            let rowText = row.innerText.toLowerCase();
            if (rowText.includes('chủ hộ')) continue; // Bỏ qua chủ hộ vì hệ thống đã tự khóa

            // Cột số 3 thường là cột Họ và tên (STT, Mã YT, Họ tên, ...)
            let nameCells = row.querySelectorAll('td');
            if(nameCells.length < 3) continue;
            let rowName = nameCells[2].innerText.trim().toLowerCase();

            // Tìm thông tin member trong danh sách đã lưu
            let member = family.members.find(m => m.name.toLowerCase() === rowName);
            if (member) {
                // Sử dụng vai trò đã được Logic chốt
                let relation = member.vaiTro || guessRelationship(family.head, member); 
                
                logText += `• ${member.name} (${member.dob || 'Không rõ NS'}): <b style="color:#d81b60">${relation}</b><br>`;
                let ngSelect = row.querySelector('ng-select');
                if (ngSelect) {
                    await selectNgOption(ngSelect, relation);
                    await sleep(200);
                }
            }
        }
        return logText;
    }

    async function processNextThemMoi() {
        if (!window.isThemMoiRunning) return;
        
        const myId = Math.random();
        window.currentThemMoiId = myId;
        
        if (window.themMoiQueue.length === 0) {
            document.getElementById('themmoi_status').innerHTML = `<span style="color: green; font-weight: bold;">🎉 ĐÃ TẠO MỚI XONG TOÀN BỘ!</span>`;
            document.getElementById('btn_start_themmoi').style.display = 'block';
            document.getElementById('btn_approve_themmoi').style.display = 'none';
            document.getElementById('btn_skip_themmoi').style.display = 'none';
            document.getElementById('btn_stop_themmoi').disabled = true;
            window.isThemMoiRunning = false;
            return;
        }

        const person = window.themMoiQueue[0];
        updateUIThemMoi();
        setButtonsDisabledThemMoi(true);
        let statusEl = document.getElementById('themmoi_status');

        if (!window.location.hash.includes('/dan-so/ho-gia-dinh/nhan-khau/nhan-khau-them-moi')) {
            statusEl.innerHTML = "🔄 Đang mở Form tạo mới...";
            window.location.hash = '/dan-so/ho-gia-dinh/nhan-khau/nhan-khau-them-moi';
            await sleep(300); 
            let waitLoading = 0;
            while (isWebLoading()) { await sleep(150); waitLoading += 150; if (waitLoading > 15000) break; } 
            await sleep(150); 
        }

        if (window.currentThemMoiId !== myId || !window.isThemMoiRunning) return;

        statusEl.innerHTML = `✍️ Đang điền dữ liệu cho: ${person.name || person.cccd}...`;
        
        let age = null;
        if (person.dob) age = calculateAge(person.dob);
        let groupValue = xacDinhNhomDoiTuong(age, ""); 

        if (person.name) fillTextInputByControlName('fullname', person.name);
        if (person.cccd) fillTextInputByControlName('citizenIdentification', person.cccd);
        if (person.phone) fillTextInputByControlName('personalPhoneNumber', person.phone);
        if (person.dob) fillDateByLabel('Ngày tháng năm sinh', person.dob);
        if (person.gender) chonGioiTinh(person.gender);

        let tinhToFill = person.parsedAddress?.tinh || 'Thành phố Hà Nội';
        let xaToFill = person.parsedAddress?.xa || window.textDienNhanhXa;
        let thonToFill = person.parsedAddress?.thon || window.textDienNhanhThon;
        let diaChiChiTiet = person.parsedAddress?.full || `Thôn ${thonToFill}, ${xaToFill}, ${tinhToFill}`;
        let thonToSearch = cleanThonXom(thonToFill);

        if (window.currentThemMoiId !== myId) return;

        await fillAngularDropdown('Tỉnh/Thành phố', tinhToFill, 0); await sleep(100); 
        await fillAngularDropdown('Xã/Phường', xaToFill, 0); await sleep(100); 
        await fillAngularDropdown('Tỉnh/Thành phố', tinhToFill, 1); await sleep(100); 
        await fillAngularDropdown('Xã/Phường', xaToFill, 1); await sleep(100); 
        if (thonToSearch) { await fillAngularDropdown('Thôn/Xóm', thonToSearch, 0); await sleep(100); } 
        if (groupValue) { await fillAngularDropdown('Nhóm đối tượng quản lý', groupValue, 0); }

        fillTextInputByControlName('currentAddress', diaChiChiTiet);
        setTimeout(() => { clickCheckboxByText('Chọn địa chỉ hiện tại là hộ khẩu thường trú'); }, 150); 

        if (window.currentThemMoiId !== myId) return; 

        statusEl.innerHTML = `🛑 ĐÃ ĐIỀN XONG. HÃY KIỂM TRA VÀ BẤM [DUYỆT & LƯU]`;
        setButtonsDisabledThemMoi(false);
    }

    async function dienThongTinVaoForm(person) {
        person.changes = []; 
        if (person && person.name) {
            let currentName = document.querySelector('input[formcontrolname="fullname"]')?.value || "";
            if (currentName && currentName.toLowerCase().trim() !== person.name.toLowerCase().trim()) {
                fillTextInputByControlName('fullname', person.name);
                person.changes.push(`Tên: ${currentName} ➡️ ${person.name}`);
            }
        }
        if (person && person.dob) {
            let currentDob = getInputValueByLabel('Ngày tháng năm sinh') || "";
            if (currentDob && currentDob !== person.dob) {
                fillDateByLabel('Ngày tháng năm sinh', person.dob);
                person.changes.push(`NS: ${currentDob} ➡️ ${person.dob}`);
            }
        }
        if (person && person.gender) {
            let isDiff = false; let currentGenderText = getCurrentGender();
            if (currentGenderText) { if (currentGenderText.toLowerCase() !== person.gender.toLowerCase()) isDiff = true; } else { isDiff = true; }
            if (isDiff) { chonGioiTinh(person.gender); person.changes.push(`GT: ${currentGenderText || 'Trống'} ➡️ ${person.gender}`); }
        }
        if (person && person.cccd) {
            let currentCccd = document.querySelector('input[formcontrolname="citizenIdentification"]')?.value || "";
            if (currentCccd !== person.cccd) {
                fillTextInputByControlName('citizenIdentification', person.cccd);
                person.changes.push(`CCCD: ${currentCccd || 'Trống'} ➡️ ${person.cccd}`);
            }
        }
        if (person && person.phone) {
            let currentPhone = document.querySelector('input[formcontrolname="personalPhoneNumber"]')?.value || "";
            if (currentPhone !== person.phone) fillTextInputByControlName('personalPhoneNumber', person.phone);
        }
        
        const xaDefault = window.textDienNhanhXa;
        const thonDefault = window.textDienNhanhThon;
        const bhytCode = getBHYTCode();
        let finalDob = person.dob || getInputValueByLabel('Ngày tháng năm sinh');
        let age = null; if (finalDob) age = calculateAge(finalDob);
        let groupValue = xacDinhNhomDoiTuong(age, bhytCode);

        if (person) { person.nhomDoiTuong = groupValue || "Không xác định"; person.maBHYT = bhytCode || ""; }

        let tinhToFill = person.parsedAddress?.tinh || 'Thành phố Hà Nội';
        let xaToFill = person.parsedAddress?.xa || xaDefault;
        let thonToFill = person.parsedAddress?.thon || thonDefault;
        let diaChiChiTiet = person.parsedAddress?.full || `Thôn ${thonToFill}, ${xaToFill}, ${tinhToFill}`;
        let thonToSearch = cleanThonXom(thonToFill);

        await fillAngularDropdown('Tỉnh/Thành phố', tinhToFill, 0); await sleep(100); 
        await fillAngularDropdown('Xã/Phường', xaToFill, 0); await sleep(100); 
        await fillAngularDropdown('Tỉnh/Thành phố', tinhToFill, 1); await sleep(100); 
        await fillAngularDropdown('Xã/Phường', xaToFill, 1); await sleep(100); 
        if (thonToSearch) { await fillAngularDropdown('Thôn/Xóm', thonToSearch, 0); await sleep(100); } 
        if (groupValue) { await fillAngularDropdown('Nhóm đối tượng quản lý', groupValue, 0); }

        fillDateByLabel('Ngày chuẩn hóa', getTodayString());
        fillTextInputByControlName('currentAddress', diaChiChiTiet);
        setTimeout(() => { clickCheckboxByText('Chọn địa chỉ hiện tại là hộ khẩu thường trú'); }, 150); 
    }

    async function fillAngularDropdown(labelText, valueToSelect, labelIndex = 0) {
        if (!valueToSelect) return false;
        return new Promise(async (resolve) => {
            const xpath = `//*[contains(text(), '${labelText}')]`;
            const matchingLabels = document.evaluate(xpath, document, null, XPathResult.ORDERED_NODE_SNAPSHOT_TYPE, null);
            let validMatches = [];
            for (let i = 0; i < matchingLabels.snapshotLength; i++) {
                let el = matchingLabels.snapshotItem(i);
                if ((el.children.length === 0 || el.tagName === 'LABEL') && el.offsetParent !== null) validMatches.push(el);
            }
            if (validMatches.length <= labelIndex) return resolve(false);

            let currentEl = validMatches[labelIndex];
            let ngSelect = null;
            for (let i = 0; i < 5; i++) { 
                if (!currentEl.parentElement) break;
                currentEl = currentEl.parentElement;
                ngSelect = currentEl.querySelector('ng-select');
                if (ngSelect) break;
            }
            if (!ngSelect) return resolve(false);

            const currentValueEl = ngSelect.querySelector('.ng-value-label');
            if (currentValueEl && currentValueEl.innerText.toLowerCase().includes(valueToSelect.toLowerCase())) return resolve(true);

            const input = ngSelect.querySelector('input[type="text"]');
            if (!input) return resolve(false);

            ngSelect.scrollIntoView({ behavior: 'smooth', block: 'center' }); await sleep(300);
            
            const container = ngSelect.querySelector('.ng-select-container');
            if (container) { container.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true })); await sleep(300); }

            input.focus(); input.value = valueToSelect; input.dispatchEvent(new Event('input', { bubbles: true }));
            
            let waitDropdown = 0;
            while(!document.querySelector('ng-dropdown-panel .ng-option') && waitDropdown < 2000) { await sleep(200); waitDropdown += 200; }

            const dropdownPanel = document.querySelector('ng-dropdown-panel');
            let clicked = false;
            if (dropdownPanel) {
                const options = Array.from(dropdownPanel.querySelectorAll('.ng-option')).filter(opt => opt.offsetParent !== null);
                let optionToClick = options.find(opt => opt.innerText.toLowerCase() === valueToSelect.toLowerCase());
                if (!optionToClick) optionToClick = options.find(opt => opt.innerText.toLowerCase().includes(valueToSelect.toLowerCase()));
                if (optionToClick) { optionToClick.click(); clicked = true; }
            }
            if (!clicked) {
                input.value = ''; input.dispatchEvent(new Event('input', { bubbles: true }));
                input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', keyCode: 27, bubbles: true })); input.blur();
            } else { input.blur(); }
            await sleep(300); resolve(clicked);
        });
    }

    async function selectNgOption(ngSelectElement, valueToSelect) {
        if (!ngSelectElement) return false;
        const input = ngSelectElement.querySelector('input[type="text"]');
        if (!input) return false;
        ngSelectElement.scrollIntoView({ behavior: 'smooth', block: 'center' }); await sleep(300);
        const container = ngSelectElement.querySelector('.ng-select-container');
        if (container) { container.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true })); await sleep(500); }

        let clicked = false;
        let dropdownPanel = document.querySelector('ng-dropdown-panel');
        if (dropdownPanel) {
            let options = Array.from(document.querySelectorAll('.ng-option')).filter(opt => opt.offsetParent !== null);
            let opt = options.find(o => o.innerText.trim().toLowerCase() === valueToSelect.toLowerCase());
            if (!opt) opt = options.find(o => o.innerText.toLowerCase().includes(valueToSelect.toLowerCase()));
            if (opt) { opt.click(); clicked = true; }
        }
        if (!clicked) {
            input.focus(); input.value = valueToSelect; input.dispatchEvent(new Event('input', { bubbles: true }));
            let waitDropdown = 0;
            while(waitDropdown < 2000) { await sleep(200); waitDropdown += 200; dropdownPanel = document.querySelector('ng-dropdown-panel'); if(dropdownPanel && !dropdownPanel.innerText.includes('Loading')) break; }
            if(dropdownPanel){
                let options = Array.from(document.querySelectorAll('ng-dropdown-panel .ng-option'));
                let opt = options.find(o => o.innerText.trim().toLowerCase() === valueToSelect.toLowerCase());
                if (!opt) opt = options.find(o => o.innerText.toLowerCase().includes(valueToSelect.toLowerCase()));
                if (opt) { opt.click(); clicked = true; }
            }
        }
        input.value = ''; input.dispatchEvent(new Event('input', { bubbles: true }));
        input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', keyCode: 27, bubbles: true })); input.blur();
        await sleep(400); return clicked;
    }

    function xacDinhNhomDoiTuong(age, bhytCode) {
        if (age !== null) { if (age <= 5) return 'Trẻ em dưới 5 tuổi'; if (age >= 60) return 'hưu trí'; }
        if (bhytCode && !bhytCode.toUpperCase().startsWith('BT') && bhytCode.length >= 2) {
            const prefix = bhytCode.substring(0, 2).toUpperCase();
            if (prefix === 'TE') return 'Trẻ em dưới 5 tuổi';
            if (prefix === 'HS') { if (age !== null && age > 18 && age <= 22) return 'Sinh viên (ĐH'; if (age !== null && age > 22) return 'lao động tự do'; return 'Học sinh (Tiểu học'; }
            if (prefix === 'SV') return 'Sinh viên (ĐH'; 
            if (prefix === 'HT') return 'hưu trí'; 
            if (['CH', 'DN', 'HC', 'CQ', 'XN', 'HX'].includes(prefix)) return 'Cán bộ, công chức'; 
            return 'lao động tự do'; 
        }
        if (age !== null) {
            if (age > 5 && age <= 18) return 'Học sinh (Tiểu học';
            if (age > 18 && age <= 22) return 'Sinh viên (ĐH';
            if (age > 22 && age < 60) return 'lao động tự do';
        }
        return 'lao động tự do'; 
    }

    function cleanThonXom(thonStr) {
        if (!thonStr) return "";
        return thonStr.replace(/^(thôn|xóm|tổ|đội|khu phố|khu|ấp|bản|phố)\s+/i, '').trim().toLowerCase();
    }

    function parseDateCustom(dateStr) {
        if(!dateStr) return "";
        let parts = dateStr.match(/\d+/g);
        if(!parts || parts.length < 3) return dateStr;
        let p1 = parseInt(parts[0]); let p2 = parseInt(parts[1]); let p3 = parseInt(parts[2]);
        if (p1 > 1000) return `${String(p3).padStart(2,'0')}/${String(p2).padStart(2,'0')}/${p1}`;
        if (p1 > 12) return `${String(p1).padStart(2,'0')}/${String(p2).padStart(2,'0')}/${p3}`;
        if (p2 > 12) return `${String(p2).padStart(2,'0')}/${String(p1).padStart(2,'0')}/${p3}`;
        return `${String(p1).padStart(2,'0')}/${String(p2).padStart(2,'0')}/${p3}`;
    }

    // TỐI ƯU HOÁ CHỨC NĂNG PHÂN TÍCH ĐỊA CHỈ
    function parseAddressString(fullAddress) {
        let result = { tinh: '', xa: '', thon: '', full: fullAddress };
        if (!fullAddress) return result;
        
        let parts = fullAddress.split(',').map(s => s.trim());
        if (parts.length === 1 && fullAddress.includes('-')) {
            parts = fullAddress.split('-').map(s => s.trim());
        }
        
        if (parts.length >= 3) {
            result.tinh = parts[parts.length - 1]; 
            result.xa = parts[parts.length - 2]; 
            
            // Cải tiến: Luôn lấy phần tử thứ 3 từ dưới lên làm Thôn/Xóm thay vì gom tất cả.
            // Giải quyết triệt để lỗi khi người dùng nhập chuỗi 4 phần 
            // VD: "Số nhà 5, Thôn Năm Trại, Xã Quốc Oai, TP Hà Nội" -> parts[1] là "Thôn Năm Trại"
            result.thon = parts[parts.length - 3]; 
            
        } else if (parts.length === 2) { 
            result.tinh = parts[1]; 
            result.xa = parts[0]; 
        }
        
        return result;
    }

    function parseExcelRow(line) {
        let parts = line.split('\t').map(s => s.trim()).filter(s => s);
        let person = { name: '', dob: '', cccd: '', gender: '', phone: '', address: '', parsedAddress: null };
        parts.forEach(p => {
            if (/^(Nam|Nữ)$/i.test(p)) person.gender = p;
            else if (/^\d{1,2}[\/\-]\d{1,2}[\/\-]\d{4}$/.test(p)) person.dob = parseDateCustom(p);
            else if (/^0\d{9}$/.test(p) && !person.phone) person.phone = p;
            else if (/^\d{9,15}$/.test(p) && !person.cccd && p.length !== 10) person.cccd = p;
            else if (p.includes(',') || p.toLowerCase().includes('xã') || p.toLowerCase().includes('phường') || p.toLowerCase().includes('thành phố')) person.address = p;
            else if (!person.name && p.includes(' ') && !/\d/.test(p) && p.length > 5 && p.length < 40) person.name = p;
        });
        person.parsedAddress = parseAddressString(person.address);
        return person;
    }

    function updateUIThemMoi() {
        const queueCountEl = document.getElementById('themmoi_queue_count');
        const queueTotalEl = document.getElementById('themmoi_total_count');
        if (queueCountEl) queueCountEl.innerText = window.totalThemMoi > 0 ? (window.totalThemMoi - window.themMoiQueue.length) : 0;
        if (queueTotalEl) queueTotalEl.innerText = window.totalThemMoi;
        
        const currentPersonEl = document.getElementById('themmoi_current_person');
        if (currentPersonEl) {
            if (window.themMoiQueue.length > 0) {
                const p = window.themMoiQueue[0];
                let htmlStr = `Tên: <span style="font-size:15px; font-weight:bold; color:#6f42c1;">${p.name || 'Không rõ'}</span>`;
                if(p.cccd) htmlStr += `<br>CCCD: ${p.cccd}`;
                if(p.dob) {
                    let age = calculateAge(p.dob); htmlStr += ` | NS: ${p.dob} (${age} tuổi)`;
                    let groupValue = xacDinhNhomDoiTuong(age, ""); if (groupValue) htmlStr += `<br>📋 Nhóm ĐT: <b>${groupValue}</b>`;
                }
                if(p.gender) htmlStr += ` | 🚻 ${p.gender}`;
                if(p.phone) htmlStr += `<br>SĐT: ${p.phone}`;
                if(p.parsedAddress && p.parsedAddress.full) htmlStr += `<br>🏡 ${p.parsedAddress.full}`;
                currentPersonEl.innerHTML = htmlStr;
            } else {
                currentPersonEl.innerText = "Trống";
            }
        }
    }

    function setButtonsDisabledThemMoi(isDisabled) {
        const btnApprove = document.getElementById('btn_approve_themmoi');
        if(btnApprove) { btnApprove.disabled = isDisabled; btnApprove.style.opacity = isDisabled ? "0.5" : "1"; }
    }

    function getCurrentGender() {
        const labels = document.querySelectorAll('label.container-radio-gender');
        for (let el of labels) {
            const input = el.querySelector('input[type="radio"]');
            if (input && input.checked) return el.innerText.trim();
        }
        return "";
    }

    function isWebLoading() {
        const xpath = `//*[normalize-space(text())='Loading...']`;
        const loadingTextNodes = document.evaluate(xpath, document, null, XPathResult.ORDERED_NODE_SNAPSHOT_TYPE, null);
        for (let i = 0; i < loadingTextNodes.snapshotLength; i++) {
            let el = loadingTextNodes.snapshotItem(i);
            if (el.tagName !== 'SCRIPT' && el.tagName !== 'STYLE' && el.offsetParent !== null) return true;
        }
        const spinnerOverlay = document.querySelector('ngx-spinner > div');
        if (spinnerOverlay && spinnerOverlay.offsetParent !== null) {
            const style = window.getComputedStyle(spinnerOverlay);
            if (style.display !== 'none' && style.opacity !== '0' && style.visibility !== 'hidden') return true;
        }
        const appLoading = document.querySelector('.app-loading');
        if (appLoading && appLoading.offsetParent !== null) {
            const style = window.getComputedStyle(appLoading);
            if (style.display !== 'none' && style.opacity !== '0') return true;
        }
        return false;
    }

    function isDuplicatePopupVisible() {
        const xpath = `//*[contains(text(), 'Xác nhận lưu nhân khẩu trùng')]`;
        const matchingLabels = document.evaluate(xpath, document, null, XPathResult.ORDERED_NODE_SNAPSHOT_TYPE, null);
        for (let i = 0; i < matchingLabels.snapshotLength; i++) {
            if (matchingLabels.snapshotItem(i).offsetParent !== null) return true;
        }
        return false;
    }

    function updateUI() {
        const queueCountEl = document.getElementById('hssk_queue_count');
        const queueTotalEl = document.getElementById('hssk_total_count');
        if (queueCountEl) queueCountEl.innerText = window.hsskBotQueue.length;
        if (queueTotalEl) queueTotalEl.innerText = window.totalChuanHoa;
        
        const currentPersonEl = document.getElementById('hssk_current_person');
        if (currentPersonEl) {
            if (window.hsskBotQueue.length > 0) {
                const p = window.hsskBotQueue[0];
                let htmlStr = `Tên: <span style="font-size:15px; font-weight:bold; color:#ff9800;">${p.name || p.hoTenTraCuu || 'Không rõ'}</span>`;
                if(p.cccd) htmlStr += `<br>CCCD: ${p.cccd}`;
                let dobToShow = p.dob || p.ngaySinhTraCuu;
                if(dobToShow) { let age = calculateAge(dobToShow); htmlStr += ` | NS: ${dobToShow} (${age} tuổi)`; }
                let genderToShow = p.gender || p.gioiTinhTraCuu;
                if(genderToShow) htmlStr += ` | 🚻 ${genderToShow}`;
                if (p.nhomDoiTuong) htmlStr += `<br>📋 Nhóm ĐT: <b>${p.nhomDoiTuong}</b>`;
                if (p.maBHYT) htmlStr += ` | BHYT: ${p.maBHYT}`;

                if (p.changes && p.changes.length > 0) {
                    htmlStr += `<br><div style="margin-top:4px; padding:4px; background:#ffeeba; border-left:3px solid #d32f2f; color:#d32f2f; font-size:11px; border-radius:3px;">⚠️ Đã sửa: <b>${p.changes.join(' | ')}</b></div>`;
                } else if (p.changes && p.changes.length === 0) {
                    htmlStr += `<br><div style="margin-top:4px; padding:4px; background:#e8f5e9; border-left:3px solid #28a745; color:#28a745; font-size:11px; border-radius:3px;">✅ Khớp 100%, không cần sửa thông tin cơ bản.</div>`;
                }
                currentPersonEl.innerHTML = htmlStr;
            } else {
                currentPersonEl.innerText = "Trống";
            }
        }
    }

    function updateStatus(text) {
        const statusEl = document.getElementById('hssk_status');
        if(statusEl) statusEl.innerText = `Trạng thái: ${text}`;
    }

    function setButtonsDisabled(isDisabled) {
        const btnApprove = document.getElementById('btn_approve');
        const btnContinue = document.getElementById('btn_continue_after_popup');
        if(btnApprove) { btnApprove.disabled = isDisabled; btnApprove.style.opacity = isDisabled ? "0.5" : "1"; }
        if(btnContinue) { btnContinue.disabled = isDisabled; btnContinue.style.opacity = isDisabled ? "0.5" : "1"; }
    }

    function findButtonByText(text) {
        const spans = Array.from(document.querySelectorAll('span.text, span.menu-name, button')).filter(s => s.offsetParent !== null);
        const span = spans.find(s => s.innerText.trim() === text);
        return span ? span.closest('button') : null;
    }

    function getTodayString() {
        const today = new Date();
        const dd = String(today.getDate()).padStart(2, '0');
        const mm = String(today.getMonth() + 1).padStart(2, '0');
        return `${dd}/${mm}/${today.getFullYear()}`;
    }

    function fillDateByLabel(labelText, dateString) {
        const xpath = `//*[contains(text(), '${labelText}')]`;
        const matchingLabels = document.evaluate(xpath, document, null, XPathResult.ORDERED_NODE_SNAPSHOT_TYPE, null);
        for (let i = 0; i < matchingLabels.snapshotLength; i++) {
            let el = matchingLabels.snapshotItem(i);
            if (el.children.length > 2 && el.tagName !== 'LABEL') continue;
            if (el.offsetParent === null) continue;
            let currentEl = el;
            let input = null;
            for (let j = 0; j < 5; j++) {
                if (!currentEl.parentElement) break;
                currentEl = currentEl.parentElement;
                input = currentEl.querySelector('input.cus-datepicker, input[bsdatepicker]');
                if (input) break;
            }
            if (input) {
                input.focus(); input.value = dateString;
                input.dispatchEvent(new Event('input', { bubbles: true }));
                input.dispatchEvent(new Event('change', { bubbles: true }));
                input.dispatchEvent(new Event('blur', { bubbles: true }));
                return;
            }
        }
    }

    function fillTextInputByControlName(controlName, value) {
        const inputs = Array.from(document.querySelectorAll(`input[formcontrolname="${controlName}"], textarea[formcontrolname="${controlName}"]`)).filter(i => i.offsetParent !== null);
        if (inputs.length > 0) {
            inputs[0].focus(); inputs[0].value = value;
            inputs[0].dispatchEvent(new Event('input', { bubbles: true }));
            inputs[0].dispatchEvent(new Event('change', { bubbles: true }));
            inputs[0].dispatchEvent(new Event('blur', { bubbles: true }));
            return true;
        }
        return false;
    }

    function clickCheckboxByText(text) {
        const xpath = `//label[contains(., '${text}')]`;
        const matchingLabels = document.evaluate(xpath, document, null, XPathResult.ORDERED_NODE_SNAPSHOT_TYPE, null);
        for (let i = 0; i < matchingLabels.snapshotLength; i++) {
            const labelNode = matchingLabels.snapshotItem(i);
            if (labelNode.offsetParent === null) continue; 
            const checkbox = labelNode.querySelector('input[type="checkbox"]');
            if (checkbox) {
                if (!checkbox.checked) {
                    const checkmark = labelNode.querySelector('.checkmark');
                    if (checkmark) checkmark.click(); else checkbox.click();
                }
                return;
            } else {
                labelNode.click(); return;
            }
        }
    }

    function getBHYTCode() {
        const elements = document.querySelectorAll('td span, td div, td');
        const bhytRegex = /^[A-Z]{2}\d{13}$/i; 
        for (let el of elements) {
            if (el.offsetParent === null) continue;
            const text = el.innerText.trim();
            if (bhytRegex.test(text)) return text;
        }
        return null;
    }

    function getInputValueByLabel(labelText) {
        const xpath = `//*[contains(text(), '${labelText}')]`;
        const matchingLabels = document.evaluate(xpath, document, null, XPathResult.ORDERED_NODE_SNAPSHOT_TYPE, null);
        for (let i = 0; i < matchingLabels.snapshotLength; i++) {
            let el = matchingLabels.snapshotItem(i);
            if (el.tagName !== 'LABEL' && el.children.length > 2) continue;
            if (el.offsetParent === null) continue;
            let currentEl = el;
            for (let j = 0; j < 5; j++) {
                if (!currentEl.parentElement) break;
                currentEl = currentEl.parentElement;
                const input = currentEl.querySelector('input');
                if (input && input.value) return input.value;
            }
        }
        return null;
    }

    function guessRelationship(head, member) {
        if (!head.dob || !member.dob) return "Khác";
        let ageHead = calculateAge(head.dob);
        let ageMember = calculateAge(member.dob);
        if (ageHead === null || ageMember === null) return "Khác";

        let diff = ageHead - ageMember; // Dương: Chủ hộ lớn tuổi hơn, Âm: Chủ hộ nhỏ tuổi hơn
        let mGender = member.gender ? member.gender.toLowerCase() : "";
        let hGender = head.gender ? head.gender.toLowerCase() : "";

        // Ưu tiên 1: Xét vợ chồng (Khác giới tính, lệch nhau tối đa 18 tuổi)
        if (Math.abs(diff) <= 18 && hGender !== "" && mGender !== "" && hGender !== mGender) {
            return mGender === "nữ" ? "Vợ" : "Chồng";
        }

        // Ưu tiên 2: Xét vai vế bề trên (Ông, Bà, Bố, Mẹ) - Chủ hộ nhỏ tuổi hơn nhiều
        if (diff <= -46) return mGender === "nam" ? "Ông" : "Bà";
        if (diff <= -16 && diff >= -45) return mGender === "nam" ? "Bố" : "Mẹ";

        // Ưu tiên 3: Xét vai vế bề dưới (Con, Cháu) - Chủ hộ lớn tuổi hơn nhiều
        if (diff >= 46) return "Cháu";
        if (diff >= 16 && diff <= 45) return mGender === "nam" ? "Con trai" : "Con gái";

        // Ưu tiên 4: Xét vai vế anh chị em (Lệch từ 1 đến 15 tuổi, cùng giới hoặc không đủ đk làm vợ chồng)
        if (diff > 0 && diff <= 15) return "Em"; 
        if (diff < 0 && diff >= -15) return mGender === "nam" ? "Anh" : "Chị";

        // Nếu bằng tuổi (diff = 0) mà cùng giới tính
        if (diff === 0) return "Khác";

        return "Khác";
    }

    function chonGioiTinh(gioiTinhText) {
        if (!gioiTinhText) return;
        const elements = document.querySelectorAll('label, span, div');
        for (let el of elements) {
            if (el.offsetParent === null) continue;
            if (el.innerText && el.innerText.trim().toLowerCase() === gioiTinhText.toLowerCase()) {
                let parent = el.parentElement;
                let isRadio = false;
                for(let i = 0; i < 3; i++) {
                    if(parent && (parent.querySelector('input[type="radio"]') || parent.tagName === 'MAT-RADIO-BUTTON' || parent.className.includes('radio'))) {
                        isRadio = true; break;
                    }
                    if(parent) parent = parent.parentElement;
                }
                if (isRadio || el.tagName === 'LABEL') {
                    el.click(); return;
                }
            }
        }
    }

    async function handleDuplicatePopup(person, dialog) {
        const text = dialog.innerText;
        
        const nameMatch = text.match(/Họ tên:\s*([^\n]+)/i);
        const dobMatch = text.match(/Ngày sinh:\s*([\d\/]+)/i);
        const genderMatch = text.match(/Giới tính:\s*(Nam|Nữ)/i);
        const cccdMatch = text.match(/(?:CCCD|định danh công dân).*?(\d{9,15})/i);

        const popName = nameMatch ? nameMatch[1].trim().toLowerCase() : "";
        const popDob = dobMatch ? dobMatch[1].trim() : "";
        const popGender = genderMatch ? genderMatch[1].trim().toLowerCase() : "";
        const popCccd = cccdMatch ? cccdMatch[1] : "";

        const exName = (person.name || "").trim().toLowerCase();
        const exDob = (person.dob || "").trim();
        const exGender = (person.gender || "").trim().toLowerCase();
        const exCccd = (person.cccd || "").trim();

        let isBasicMatch = true;
        let mismatchDetails = [];

        if (exName && popName && exName !== popName) { isBasicMatch = false; mismatchDetails.push(`Tên`); }
        if (exDob && popDob && exDob !== popDob) { isBasicMatch = false; mismatchDetails.push(`NS`); }
        if (exGender && popGender && exGender !== popGender) { isBasicMatch = false; mismatchDetails.push(`GT`); }

        if (!isBasicMatch) {
             return { handled: false, isCccdMatched: false, reason: `Lệch: ${mismatchDetails.join(', ')}` };
        }

        let isCccdMatched = false;
        if (exCccd && popCccd) {
            if (exCccd === popCccd) {
                isCccdMatched = true; 
            } else {
                return { handled: false, isCccdMatched: false, reason: `Lệch CCCD (Web: ${popCccd})` }; 
            }
        } else if (popCccd && !exCccd) {
             isCccdMatched = true; 
        } else {
             isCccdMatched = false; 
        }

        const spans = Array.from(dialog.querySelectorAll('span'));
        const linkSpan = spans.find(s => s.innerText.trim().toLowerCase() === 'vào đây');
        if (linkSpan) {
            linkSpan.click();
            return { handled: true, isCccdMatched: isCccdMatched, reason: '' };
        }
        return { handled: false, isCccdMatched: false, reason: 'Lỗi không tìm thấy nút "vào đây"' };
    }

})();
