// 會議室對應的電子郵件地址
const ROOM_EMAILS = {
    '多功能會議室': 'c_1883f40drcs8kh6lkvl36oh7qlhq6@resource.calendar.google.com',
    '明德堂大樓-3樓-大禮堂': 'c_1884dls2ka6oqjkvg67tvv4ob9jgu@resource.calendar.google.com',
    '明德堂大樓-6樓-校史館': 'c_1882ckhmb7kv6i51mub5ji2a6rsos@resource.calendar.google.com',
    '明德堂大樓-6樓-圖書館': 'c_18855aua3r1j6jg3l5ofbdl4d1deu@resource.calendar.google.com',
    '餐飲觀光大樓-4樓-階梯會議室': 'c_18803q99rlfdoiqnn64otgugts66m@resource.calendar.google.com'
};

// 會議室對應的公開網址
const ROOM_PUBLIC_URLS = {
    '多功能會議室': 'https://calendar.google.com/calendar/embed?src=c_1883f40drcs8kh6lkvl36oh7qlhq6%40resource.calendar.google.com&ctz=Asia%2FTaipei',
    '明德堂大樓-3樓-大禮堂': 'https://calendar.google.com/calendar/embed?src=c_1884dls2ka6oqjkvg67tvv4ob9jgu%40resource.calendar.google.com&ctz=Asia%2FTaipei',
    '明德堂大樓-6樓-校史館': 'https://calendar.google.com/calendar/embed?src=c_1882ckhmb7kv6i51mub5ji2a6rsos%40resource.calendar.google.com&ctz=Asia%2FTaipei',
    '明德堂大樓-6樓-圖書館': 'https://calendar.google.com/calendar/embed?src=c_18855aua3r1j6jg3l5ofbdl4d1deu%40resource.calendar.google.com&ctz=Asia%2FTaipei',
    '餐飲觀光大樓-4樓-階梯會議室': 'https://calendar.google.com/calendar/embed?src=c_18803q99rlfdoiqnn64otgugts66m%40resource.calendar.google.com&ctz=Asia%2FTaipei'
};

// DOM 元素
const elements = {
    department: document.getElementById('department'),
    eventName: document.getElementById('eventName'),
    roomSelect: document.getElementById('roomSelect'),
    eventDate: document.getElementById('eventDate'),
    startTime: document.getElementById('startTime'),
    endTime: document.getElementById('endTime'),
    generateBtn: document.getElementById('generateBtn')
};

// 日期選取器實例
let datePicker = null;

// 時間限制常數
const BUSINESS_MIN_TIME = '06:00';
const BUSINESS_MAX_TIME = '22:00';

// 當前生成的連結
let currentLink = '';

// 初始化
document.addEventListener('DOMContentLoaded', function() {
    initializeForm();
    setupEventListeners();
    handleUrlHash();
    initializeTimePickers();
    initializeDatePicker();
});

// 初始化表單
function initializeForm() {
    // 不預設任何時間，讓使用者自行選擇
}

// 處理 URL 片段（hash）
function handleUrlHash() {
    const hash = window.location.hash.substring(1); // 移除 # 符號
    
    if (hash) {
        // 解碼 URL 編碼的字符
        const decodedHash = decodeURIComponent(hash);
        
        // 檢查是否為有效的會議室名稱
        const roomNames = Object.keys(ROOM_EMAILS);
        const matchedRoom = roomNames.find(room => room === decodedHash);
        
        if (matchedRoom) {
            // 預選會議室
            elements.roomSelect.value = matchedRoom;
            showNotification(`已預選會議室：${matchedRoom}`, 'success');
            
            // 觸發表單驗證
            validateForm();
        } else {
            // 顯示錯誤訊息
            showNotification(`找不到會議室：${decodedHash}`, 'error');
        }
    }
}

// 設置事件監聽器
function setupEventListeners() {
    elements.generateBtn.addEventListener('click', generateCalendarLink);
    
    // 表單驗證
    elements.department.addEventListener('input', validateForm);
    elements.eventName.addEventListener('input', validateForm);
    elements.roomSelect.addEventListener('change', validateForm);
    elements.eventDate.addEventListener('change', onDateChange);
    elements.eventDate.addEventListener('blur', onDateChange);
    elements.eventDate.addEventListener('input', onDateInput);
    elements.startTime.addEventListener('change', onTimeChange);
    elements.endTime.addEventListener('change', onTimeChange);
    elements.startTime.addEventListener('blur', onTimeChange);
    elements.endTime.addEventListener('blur', onTimeChange);
    // 即時輸入正規化（防抖）與 Enter 快速提交
    elements.startTime.addEventListener('input', onTimeInput);
    elements.endTime.addEventListener('input', onTimeInput);
    elements.startTime.addEventListener('keydown', onTimeKeyDown);
    elements.endTime.addEventListener('keydown', onTimeKeyDown);
    
    // 鍵盤支援
    document.addEventListener('keydown', function(e) {
        if (e.key === 'Enter' && e.ctrlKey) {
            generateCalendarLink();
        }
    });
    
    // 監聽 URL 變化
    window.addEventListener('hashchange', handleUrlHash);
}

// 日期：支援民國年輸入（例如 114/01/02、1140102、114-1-2、114）
function onDateChange(e) {
    const normalized = normalizeRocDateString(e.target.value, elements.eventDate.value);
    if (normalized && normalized !== elements.eventDate.value) {
        elements.eventDate.value = normalized;
        if (datePicker) datePicker.setDate(normalized, true, 'Y-m-d');
    } else {
        // 若只輸入民國年，嘗試將面板跳到對應年
        const rocYear = extractPureRocYear(e.target.value);
        if (rocYear !== null && datePicker && datePicker.currentYear !== rocYear + 1911) {
            datePicker.changeYear(rocYear + 1911);
        }
    }
    // 日期變更後依據是否為今日，調整時間最小值
    const minTimeTodayAware = getMinTimeForSelectedDate();

    // 若既有時間早於最小允許值，則自動上調到最小值
    if (elements.startTime.value && compareTime(elements.startTime.value, minTimeTodayAware) < 0) {
        elements.startTime.value = minTimeTodayAware;
    }
    if (elements.endTime.value && compareTime(elements.endTime.value, minTimeTodayAware) < 0) {
        elements.endTime.value = minTimeTodayAware;
    }
    validateForm();
}

const handleDateInput = debounce((input) => {
    const normalized = normalizeRocDateString(input.value, elements.eventDate.value);
    if (normalized && normalized !== input.value) {
        input.value = normalized;
        validateForm();
    }
}, 200);

function onDateInput(e) {
    handleDateInput(e.target);
}

// 變更時間時，將時間對齊到5分鐘刻度並驗證
function onTimeChange(e) {
    const input = e.target;
    const normalized = normalizeTimeString(input.value);
    if (normalized && normalized !== input.value) {
        input.value = normalized;
    }
    const rounded = roundToFiveMinutes(input.value);
    if (rounded !== input.value) {
        input.value = rounded;
    }
    // 動態最小時間：今日則不可早於當前下一個5分鐘
    const minTime = getMinTimeForSelectedDate();
    if (elements.eventDate.value && compareTime(input.value, minTime) < 0) {
        input.value = minTime;
    }
    validateForm();
}

// 防抖工具
function debounce(fn, delay = 250) {
    let timer = null;
    return function(...args) {
        if (timer) clearTimeout(timer);
        timer = setTimeout(() => fn.apply(this, args), delay);
    };
}

const handleTimeInput = debounce((input) => {
    // 僅在純數字長度 3-4（如 830、1740）或已含冒號時嘗試正規化
    const raw = input.value.trim();
    const onlyDigits = /^\d{3,4}$/.test(raw);
    const hasColon = /^\d{1,2}:\d{0,2}$/.test(raw);
    if (!onlyDigits && !hasColon) return;
    const normalized = normalizeTimeString(raw);
    if (normalized && normalized !== input.value) {
        input.value = normalized;
        // 今日的最小時間限制
        const minTime = getMinTimeForSelectedDate();
        if (elements.eventDate.value && compareTime(input.value, minTime) < 0) {
            input.value = minTime;
        }
        // 輕量驗證，不彈出提示
        validateForm();
    }
}, 200);

function onTimeInput(e) {
    handleTimeInput(e.target);
}

function onTimeKeyDown(e) {
    if (e.key === 'Enter') {
        e.preventDefault();
        onTimeChange(e);
    }
}

// 將 HH:MM 對齊到最接近的5分鐘（四捨五入）
function roundToFiveMinutes(hhmm) {
    if (!hhmm) return hhmm;
    const [hStr, mStr] = hhmm.split(':');
    let hours = parseInt(hStr, 10);
    let minutes = parseInt(mStr, 10);
    if (Number.isNaN(hours) || Number.isNaN(minutes)) return hhmm;
    const total = hours * 60 + minutes;
    const roundedTotal = Math.round(total / 5) * 5;
    const rHours = Math.min(23, Math.floor(roundedTotal / 60));
    const rMinutes = Math.min(55, roundedTotal % 60);
    return `${String(rHours).padStart(2, '0')}:${String(rMinutes).padStart(2, '0')}`;
}

// 正規化任意輸入成 HH:MM（支援 1740 -> 17:40、830 -> 08:30、9 -> 09:00）
function normalizeTimeString(value, min = '06:00', max = '22:00') {
    if (!value) return value;
    const trimmed = String(value).trim();
    if (!trimmed) return trimmed;

    const toMinutes = (t) => {
        const [h, m] = t.split(':').map(n => parseInt(n, 10));
        return h * 60 + m;
    };
    const fromMinutes = (mins) => {
        const h = String(Math.floor(mins / 60)).padStart(2, '0');
        const m = String(mins % 60).padStart(2, '0');
        return `${h}:${m}`;
    };
    const clamp = (v, lo, hi) => Math.max(lo, Math.min(hi, v));

    // 已經是 HH:MM
    if (/^\d{1,2}:\d{1,2}$/.test(trimmed)) {
        const [h, m] = trimmed.split(':').map(x => parseInt(x, 10));
        const hm = clamp(h * 60 + m, toMinutes(min), toMinutes(max));
        return fromMinutes(hm);
    }

    // 僅數字：830、1740、9、07等
    const digits = trimmed.replace(/[^0-9]/g, '');
    if (!digits) return '';

    let hours = 0, minutes = 0;
    if (digits.length <= 2) {
        hours = parseInt(digits, 10);
        minutes = 0;
    } else {
        const mm = digits.slice(-2);
        const hh = digits.slice(0, -2);
        hours = parseInt(hh, 10);
        minutes = parseInt(mm, 10);
    }

    if (Number.isNaN(hours)) hours = 0;
    if (Number.isNaN(minutes)) minutes = 0;
    // 邏輯修正：分鐘 > 59 時進位到小時
    if (minutes > 59) {
        hours += Math.floor(minutes / 60);
        minutes = minutes % 60;
    }
    // 限制 0-23、0-59，稍後再夾在 min-max
    hours = Math.max(0, Math.min(23, hours));
    minutes = Math.max(0, Math.min(59, minutes));

    const total = hours * 60 + minutes;
    const clamped = clamp(total, toMinutes(min), toMinutes(max));
    return fromMinutes(clamped);
}

// 正規化民國年日期輸入為 YYYY-MM-DD
// 支援：
// - 114 -> 2025-<現有月>-<現有日>（若無現有值則取今天月日）
// - 114/01/02、114-1-2、1140102 -> 2025-01-02
function normalizeRocDateString(inputValue, currentValue) {
    if (!inputValue) return inputValue;
    const raw = String(inputValue).trim();
    if (!raw) return raw;

    // 判斷是否為疑似民國年開頭（2-3位數年）
    const rocPatterns = [
        /^(\d{2,3})[\/\-](\d{1,2})[\/\-](\d{1,2})$/, // 114/1/2 或 114-1-2
        /^(\d{7,8})$/ // 1140102 或 1031225（允許 7-8 位，月日可不補零）
    ];

    const now = new Date();
    const pad2 = (n) => String(n).padStart(2, '0');

    const useCurrentOrToday = () => {
        let month, day;
        if (/^\d{4}-\d{2}-\d{2}$/.test(currentValue || '')) {
            const [y, m, d] = currentValue.split('-');
            month = parseInt(m, 10);
            day = parseInt(d, 10);
        } else {
            month = now.getMonth() + 1;
            day = now.getDate();
        }
        return { month, day };
    };

    // 格式：114/01/02 或 114-1-2
    const m1 = raw.match(rocPatterns[0]);
    if (m1) {
        const roc = parseInt(m1[1], 10);
        let month = parseInt(m1[2], 10);
        let day = parseInt(m1[3], 10);
        const year = roc + 1911;
        month = Math.max(1, Math.min(12, month));
        day = Math.max(1, Math.min(31, day));
        return `${year}-${pad2(month)}-${pad2(day)}`;
    }

    // 格式：1140102 或 1031225
    const m2 = raw.match(rocPatterns[1]);
    if (m2) {
        const digits = m2[1];
        const roc = parseInt(digits.slice(0, digits.length - 4), 10);
        const md = digits.slice(-4);
        let month = parseInt(md.slice(0, 2), 10);
        let day = parseInt(md.slice(2, 4), 10);
        const year = roc + 1911;
        month = Math.max(1, Math.min(12, month));
        day = Math.max(1, Math.min(31, day));
        return `${year}-${pad2(month)}-${pad2(day)}`;
    }

    // 只有民國年：114（不自動補齊月日，避免預設日期）
    if (/^\d{2,3}$/.test(raw)) {
        return inputValue; // 保持原樣，要求使用者輸入完整日期
    }

    // 已經是 YYYY-MM-DD 或其他格式，直接回傳原值
    return inputValue;
}

// 解析「只有民國年份」的輸入，回傳數字年份或 null
function extractPureRocYear(value) {
    if (!value) return null;
    const raw = String(value).trim();
    if (/^\d{2,3}$/.test(raw)) {
        const roc = parseInt(raw, 10);
        if (!Number.isNaN(roc)) return roc;
    }
    return null;
}

// 將時間字串加減固定分鐘數並限制在允許範圍
function adjustTimeByMinutes(hhmm, deltaMinutes, min = '06:00', max = '22:00') {
    const toMinutes = (t) => {
        const [h, m] = t.split(':').map(n => parseInt(n, 10));
        return h * 60 + m;
    };
    const clamp = (v, lo, hi) => Math.max(lo, Math.min(hi, v));
    const step = 5;
    const cur = toMinutes(hhmm || min);
    const next = clamp(cur + deltaMinutes, toMinutes(min), toMinutes(max));
    const snapped = Math.round(next / step) * step;
    const h = String(Math.floor(snapped / 60)).padStart(2, '0');
    const m = String(snapped % 60).padStart(2, '0');
    return `${h}:${m}`;
}

// 取得「下一個 5 分鐘」的 HH:MM（若剛好在刻度上，則取下一格）
function getNextFiveMinutes() {
    const now = new Date();
    const minutes = now.getMinutes();
    const add = (5 - (minutes % 5)) % 5 || 5;
    const t = new Date(now.getTime());
    t.setMinutes(minutes + add, 0, 0);
    const hh = String(t.getHours()).padStart(2, '0');
    const mm = String(t.getMinutes()).padStart(2, '0');
    return `${hh}:${mm}`;
}

// 取得依據選取日期的最小時間（今日=下一個5分鐘，其他日=營業起始時間）
function getMinTimeForSelectedDate() {
    const date = elements.eventDate.value;
    const today = formatDateForInput(new Date());
    if (date && date === today) {
        const min = getNextFiveMinutes();
        if (compareTime(min, BUSINESS_MAX_TIME) > 0) return BUSINESS_MAX_TIME;
        if (compareTime(min, BUSINESS_MIN_TIME) < 0) return BUSINESS_MIN_TIME;
        return min;
    }
    return BUSINESS_MIN_TIME;
}

// 比較時間字串 HH:MM，回傳 -1/0/1
function compareTime(a, b) {
    const [ah, am] = a.split(':').map(n => parseInt(n, 10));
    const [bh, bm] = b.split(':').map(n => parseInt(n, 10));
    const av = ah * 60 + am;
    const bv = bh * 60 + bm;
    return av === bv ? 0 : (av < bv ? -1 : 1);
}

// 初始化 ClockPicker 時間選取器（24小時制，採用 https://weareoutman.github.io/clockpicker/）
function initializeTimePickers() {
    if (typeof jQuery === 'undefined' || typeof jQuery.fn.clockpicker === 'undefined') {
        console.warn('ClockPicker 或 jQuery 尚未載入');
        return;
    }

    const clockOptions = {
        placement: 'bottom',
        align: 'left',
        autoclose: true,
        donetext: '確定',
        twelvehour: false,
        vibrate: true,
        afterDone: function() {
            if (this.input && this.input[0]) {
                this.input[0].dispatchEvent(new Event('change', { bubbles: true }));
            }
            validateForm();
        },
        afterShow: function() {
            // 標註非營業時間刻度 (06:00 前與 22:00 後稍微淡化)
            $('.clockpicker-popover:visible .clockpicker-hours .clockpicker-tick').each(function() {
                const h = parseInt($(this).text(), 10);
                if (h < 6 || h > 22) {
                    $(this).addClass('out-of-business');
                }
            });
        }
    };

    $('#startTime').clockpicker({
        ...clockOptions,
        'default': '08:00'
    });

    $('#endTime').clockpicker({
        ...clockOptions,
        'default': '09:00'
    });

    // 點擊圖示開啟 ClockPicker
    const startIcon = document.getElementById('startTimeIcon');
    if (startIcon) {
        startIcon.addEventListener('click', (e) => {
            e.stopPropagation();
            $('#startTime').clockpicker('show');
        });
    }
    const endIcon = document.getElementById('endTimeIcon');
    if (endIcon) {
        endIcon.addEventListener('click', (e) => {
            e.stopPropagation();
            $('#endTime').clockpicker('show');
        });
    }
}


// 初始化日期 Flatpickr（不預設日期，允許輸入）
function initializeDatePicker() {
    if (typeof flatpickr !== 'function') return;
         datePicker = flatpickr('#eventDate', {
         dateFormat: 'Y-m-d',
         allowInput: true,
         clickOpens: true,
         defaultDate: null,
         minDate: 'today',
         locale: (window.flatpickr && window.flatpickr.l10ns && window.flatpickr.l10ns.zh_tw) ? window.flatpickr.l10ns.zh_tw : undefined,
         onChange: validateForm,
         disableMobile: true
     });
}



// 格式化日期為 input 格式 (YYYY-MM-DD)
function formatDateForInput(date) {
    return date.toISOString().split('T')[0];
}

// 格式化日期為 Google 日曆 UTC 格式 (YYYYMMDDTHHMMSSZ)
function formatDateForGoogleUTC(date, time) {
    const [year, month, day] = date.split('-');
    const [hour, minute] = time.split(':');
    // 台灣時區 +08:00
    const localDate = new Date(`${year}-${month}-${day}T${hour}:${minute}:00+08:00`);
    return localDate.toISOString().replace(/[-:]/g, '').replace(/\.\d{3}Z$/, 'Z');
}

// CORS 代理列表（併行查詢）
const CORS_PROXIES = [
    {
        name: 'gas-proxy',
        buildUrl: (targetUrl) => `https://script.google.com/macros/s/AKfycbyu3hZvE_-V5fvuBLYEogMCs2kboFRObzplU55UCQtWb-BBDgWELLC4cjr18wzHVm4l/exec?url=${encodeURIComponent(targetUrl)}`
    }
];

// 透過單一代理取得 iCal 資料（附逾時控制）
async function fetchViaProxy(proxy, targetUrl) {
    const url = proxy.buildUrl(targetUrl);
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), 10000); // 10 秒逾時
    try {
        const response = await fetch(url, {
            method: 'GET',
            headers: { 'Accept': 'text/calendar,text/plain,*/*' },
            signal: controller.signal
        });
        clearTimeout(timeout);
        if (!response.ok) {
            throw new Error(`HTTP ${response.status} ${response.statusText}`);
        }
        const text = await response.text();
        // 驗證回傳的確實是 iCal 資料
        if (!text.includes('BEGIN:VCALENDAR')) {
            throw new Error('回傳內容非有效的 iCal 格式');
        }
        console.log(`[衝突檢查] ${proxy.name} 代理成功取得 iCal 資料`);
        return { proxy: proxy.name, data: text };
    } catch (err) {
        clearTimeout(timeout);
        const errMsg = err.name === 'AbortError' ? '連線逾時' : err.message;
        console.warn(`[衝突檢查] ${proxy.name} 代理失敗: ${errMsg}`);
        throw new Error(`${proxy.name}: ${errMsg}`);
    }
}

// 併行查詢所有 CORS 代理，檢查會議室衝突
// 回傳: { hasConflict: boolean, conflicts: Array, allFailed: boolean }
async function checkRoomConflictParallel(icalUrl, startDateTime, endDateTime) {
    // 所有代理併行發送請求
    const promises = CORS_PROXIES.map(proxy =>
        fetchViaProxy(proxy, icalUrl).then(result => {
            // 成功取得資料後立即解析衝突
            const conflicts = parseICalForConflicts(result.data, startDateTime, endDateTime);
            return { proxy: result.proxy, conflicts, success: true };
        }).catch(err => {
            return { proxy: err.message, conflicts: [], success: false, error: err.message };
        })
    );

    const results = await Promise.all(promises);

    // 統計結果
    const successResults = results.filter(r => r.success);
    const conflictResults = successResults.filter(r => r.conflicts.length > 0);

    console.log(`[衝突檢查] 結果: ${successResults.length}/${results.length} 代理成功, ${conflictResults.length} 個發現衝突`);

    // 所有代理都失敗
    if (successResults.length === 0) {
        return { hasConflict: false, conflicts: [], allFailed: true };
    }

    // 只要有任何一個代理發現衝突，就判定為衝突
    if (conflictResults.length > 0) {
        // 合併所有衝突（去重）
        const allConflicts = [];
        const seen = new Set();
        for (const r of conflictResults) {
            for (const c of r.conflicts) {
                const key = `${c.start?.getTime()}-${c.end?.getTime()}-${c.summary}`;
                if (!seen.has(key)) {
                    seen.add(key);
                    allConflicts.push(c);
                }
            }
        }
        return { hasConflict: true, conflicts: allConflicts, allFailed: false };
    }

    // 至少有一個代理成功且沒有衝突
    return { hasConflict: false, conflicts: [], allFailed: false };
}

// 檢查會議室衝突（主入口）
async function checkRoomConflict(room, date, startTime, endTime) {
    try {
        const roomEmail = ROOM_EMAILS[room];
        const icalUrl = `https://calendar.google.com/calendar/ical/${roomEmail}/public/basic.ics`;
        // 解析日期時間（台灣時區 +08:00）
        const [year, month, day] = date.split('-');
        const [startHour, startMinute] = startTime.split(':');
        const [endHour, endMinute] = endTime.split(':');
        const localStart = new Date(`${year}-${month}-${day}T${startHour}:${startMinute}:00+08:00`);
        const localEnd = new Date(`${year}-${month}-${day}T${endHour}:${endMinute}:00+08:00`);
        // 直接用 UTC 時間做比對
        const startDateTime = new Date(localStart.getTime());
        const endDateTime = new Date(localEnd.getTime());
        // 併行查詢所有 CORS 代理
        const result = await checkRoomConflictParallel(icalUrl, startDateTime, endDateTime);

        if (result.allFailed) {
            // 所有代理都失敗 → 阻止跳轉（安全起見）
            showNotification('⚠️ 所有代理均無法連線，無法確認會議室可用性，請稍後再試或手動確認', 'error');
            return true; // 回傳 true 以阻止跳轉
        }

        if (result.hasConflict) {
            const conflictInfo = result.conflicts.map(c => {
                const startStr = c.start ? c.start.toLocaleString('zh-TW') : '未知時間';
                const endStr = c.end ? c.end.toLocaleString('zh-TW') : '未知時間';
                return `${c.summary || '未命名事件'} (${startStr} - ${endStr})`;
            }).join(', ');
            showNotification(`檢測到時間衝突：${conflictInfo}`, 'error');
            return true;
        }

        showNotification('會議室時間可用，無衝突 ✓', 'success');
        return false;
    } catch (error) {
        console.error('檢查會議室衝突時發生錯誤:', error);
        // 發生未預期錯誤 → 阻止跳轉（安全起見）
        showNotification('⚠️ 檢查衝突時發生錯誤，請稍後再試或手動確認會議室可用性', 'error');
        return true; // 回傳 true 以阻止跳轉
    }
}

// 處理 iCal 折行（RFC 5545 標準）
function unfoldICalLines(icalData) {
    const lines = icalData.split('\n');
    const unfolded = [];
    for (let i = 0; i < lines.length; i++) {
        if (lines[i][0] === ' ' && unfolded.length > 0) {
            unfolded[unfolded.length - 1] += lines[i].substring(1);
        } else {
            unfolded.push(lines[i]);
        }
    }
    return unfolded;
}

function parseICalForConflicts(icalData, startDateTime, endDateTime) {
    const conflicts = [];
    const lines = unfoldICalLines(icalData);
    let currentEvent = null;
    let inEvent = false;

    for (let i = 0; i < lines.length; i++) {
        let line = lines[i].trim();
        if (line === 'BEGIN:VEVENT') {
            inEvent = true;
            currentEvent = {};
        } else if (line === 'END:VEVENT') {
            inEvent = false;
            // 只比對有正確 start/end 且為有效日期的事件
            if (currentEvent && currentEvent.start instanceof Date && !isNaN(currentEvent.start) && currentEvent.end instanceof Date && !isNaN(currentEvent.end)) {
                if (isTimeConflict(currentEvent, startDateTime, endDateTime)) {
                    conflicts.push(currentEvent);
                }
            }
            currentEvent = null;
        } else if (inEvent && currentEvent) {
            const colonIndex = line.indexOf(':');
            if (colonIndex > 0) {
                const key = line.substring(0, colonIndex);
                const value = line.substring(colonIndex + 1);
                if (key === 'SUMMARY') {
                    currentEvent.summary = value;
                } else if (key.startsWith('DTSTART')) {
                    currentEvent.start = parseICalDateTime(key, value);
                } else if (key.startsWith('DTEND')) {
                    currentEvent.end = parseICalDateTime(key, value);
                }
            }
        }
    }
    return conflicts;
}

// 修正：更穩健的 iCal 日期時間解析器
function parseICalDateTime(key, value) {
    try {
        let year, month, day, hour, minute, second, isoString;

        // 全天事件: VALUE=DATE:YYYYMMDD
        if (key.includes('VALUE=DATE')) {
            year = value.substring(0, 4);
            month = value.substring(4, 6);
            day = value.substring(6, 8);
            isoString = `${year}-${month}-${day}T00:00:00Z`; // 標準化為 UTC 午夜
            return new Date(isoString);
        }

        // UTC 時間: YYYYMMDDTHHMMSSZ
        if (value.endsWith('Z')) {
            year = value.substring(0, 4);
            month = value.substring(4, 6);
            day = value.substring(6, 8);
            hour = value.substring(9, 11);
            minute = value.substring(11, 13);
            second = value.substring(13, 15);
            isoString = `${year}-${month}-${day}T${hour}:${minute}:${second}Z`;
            return new Date(isoString);
        }

        // 帶有時區標識的時間: TZID=...:YYYYMMDDTHHMMSS
        if (key.includes('TZID')) {
            const timePart = value.split(':').pop();
            year = timePart.substring(0, 4);
            month = timePart.substring(4, 6);
            day = timePart.substring(6, 8);
            hour = timePart.substring(9, 11);
            minute = timePart.substring(11, 13);
            second = timePart.substring(13, 15);
            isoString = `${year}-${month}-${day}T${hour}:${minute}:${second}`;
            // 簡化處理：假設為台灣時區。完整的解決方案需要時區庫。
            return new Date(isoString + "+08:00");
        }
    } catch (e) {
        console.error(`無法解析日期: ${key}:${value}`, e);
        return null;
    }

    return null; // 不支援的格式
}

// 檢查時間衝突
function isTimeConflict(event, startDateTime, endDateTime) {
    if (!event.start || !event.end) return false;
    // 檢查是否有重疊
    return (event.start < endDateTime && event.end > startDateTime);
}

// 驗證表單
function validateForm() {
    const department = elements.department.value.trim();
    const eventName = elements.eventName.value.trim();
    const room = elements.roomSelect.value;
    const date = elements.eventDate.value;
    const startTime = elements.startTime.value;
    const endTime = elements.endTime.value;
    
    const isValid = department && eventName && room && date && startTime && endTime;
    
    elements.generateBtn.disabled = !isValid;
    
    // 驗證時間邏輯
    if (startTime && endTime && startTime >= endTime) {
        showNotification('結束時間必須晚於開始時間', 'error');
        elements.generateBtn.disabled = true;
    }
    
    return isValid;
}

// 產生日曆連結
async function generateCalendarLink() {
    const department = elements.department.value.trim();
    const eventName = elements.eventName.value.trim();
    const room = elements.roomSelect.value;
    const date = elements.eventDate.value;
    const startTime = elements.startTime.value;
    const endTime = elements.endTime.value;
    
    if (!validateForm()) {
        showNotification('請填寫處室單位、活動名稱並選擇會議室', 'error');
        return;
    }
    
    try {
        // 顯示載入狀態
        elements.generateBtn.innerHTML = '<span class="loading"></span>檢查衝突中...';
        elements.generateBtn.disabled = true;
        
        // 檢查會議室衝突
        const hasConflict = await checkRoomConflict(room, date, startTime, endTime);
        
        if (hasConflict) {
            showNotification('檢測到會議室時間衝突，請重新選擇時間', 'error');
            elements.generateBtn.innerHTML = '產生日曆連結';
            elements.generateBtn.disabled = false;
            return;
        }
        
        // 更新載入狀態
        elements.generateBtn.innerHTML = '<span class="loading"></span>產生中...';
        
        // 獲取會議室郵箱
        const roomEmail = ROOM_EMAILS[room];
        if (!roomEmail) {
            throw new Error('找不到會議室郵箱地址');
        }
        
        // 格式化日期時間為 UTC
        const startDateTime = formatDateForGoogleUTC(date, startTime);
        const endDateTime = formatDateForGoogleUTC(date, endTime);
        
        // 構建活動標題：【處室單位】活動名稱
        const eventTitle = `【${department}】${eventName}`;
        
        // 構建 Google 日曆 URL（需要登入驗證的格式）
        const calendarUrl = buildGoogleCalendarUrl({
            text: eventTitle,
            dates: `${startDateTime}/${endDateTime}`,
            add: roomEmail
        });
        
        // 更新當前連結
        currentLink = calendarUrl;
        
        // 自動複製到剪貼簿
        await copyToClipboard();
        
        // 跳轉到當前頁面
        window.location.href = calendarUrl;
        
        showNotification('日曆連結已產生、複製並開啟', 'success');
        
    } catch (error) {
        console.error('產生連結時發生錯誤:', error);
        showNotification('產生連結時發生錯誤，請重試', 'error');
    } finally {
        // 恢復按鈕狀態
        elements.generateBtn.innerHTML = '產生日曆連結';
        elements.generateBtn.disabled = false;
    }
}

// 構建 Google 日曆 URL（需要登入驗證的格式）
function buildGoogleCalendarUrl(params) {
    // 構建 Google 日曆 eventedit URL
    const baseUrl = 'https://calendar.google.com/calendar/u/0/r/eventedit';
    
    // 對參數進行第一次編碼
    const encodedText = encodeURIComponent(params.text);
    const encodedAdd = encodeURIComponent(params.add);
    
    // 構建查詢字串
    let queryString = `?text=${encodedText}&add=${encodedAdd}`;
    
    if (params.dates) {
        queryString += `&dates=${params.dates}`;
    }
    
    // 完整的日曆 URL
    const fullCalendarUrl = `${baseUrl}${queryString}`;
    
    // 對完整的日曆 URL 進行第二次編碼（雙重編碼）
    const doubleEncodedUrl = encodeURIComponent(fullCalendarUrl);
    
    // 構建最終的帳戶選擇器 URL
    return `https://accounts.google.com/AccountChooser/signinchooser?continue=${doubleEncodedUrl}&hd=g.ypvs.tyc.edu.tw&flowName=GlifWebSignIn&flowEntry=AccountChooser`;
}

// 複製到剪貼簿
async function copyToClipboard() {
    if (!currentLink) {
        showNotification('沒有可複製的連結', 'error');
        return;
    }
    
    try {
        await navigator.clipboard.writeText(currentLink);
        return true;
    } catch (error) {
        console.error('複製失敗:', error);
        return false;
    }
}

// 顯示通知
function showNotification(message, type = 'success') {
    // 移除現有通知
    const existingNotification = document.querySelector('.notification');
    if (existingNotification) {
        existingNotification.remove();
    }
    
    // 創建新通知
    const notification = document.createElement('div');
    notification.className = `notification ${type}`;
    notification.textContent = message;
    
    document.body.appendChild(notification);
    
    // 自動移除通知
    setTimeout(() => {
        if (notification.parentNode) {
            notification.remove();
        }
    }, 4000);
}

// 導出函數供測試使用
if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        generateCalendarLink,
        ROOM_EMAILS,
        validateForm
    };
} 