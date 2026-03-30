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

// Flatpickr 實例
let startTimePicker = null;
let endTimePicker = null;
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
    if (startTimePicker) startTimePicker.set('minTime', minTimeTodayAware);
    if (endTimePicker) endTimePicker.set('minTime', minTimeTodayAware);

    // 若既有時間早於最小允許值，則自動上調到最小值
    if (elements.startTime.value && compareTime(elements.startTime.value, minTimeTodayAware) < 0) {
        elements.startTime.value = minTimeTodayAware;
        startTimePicker && startTimePicker.setDate(minTimeTodayAware, true, 'H:i');
    }
    if (elements.endTime.value && compareTime(elements.endTime.value, minTimeTodayAware) < 0) {
        elements.endTime.value = minTimeTodayAware;
        endTimePicker && endTimePicker.setDate(minTimeTodayAware, true, 'H:i');
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
    // 同步 Flatpickr 顯示
    if (input === elements.startTime && typeof startTimePicker?.setDate === 'function') {
        startTimePicker.set('minTime', minTime);
        startTimePicker.setDate(input.value, true, 'H:i');
    } else if (input === elements.endTime && typeof endTimePicker?.setDate === 'function') {
        endTimePicker.set('minTime', minTime);
        endTimePicker.setDate(input.value, true, 'H:i');
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
        // 同步 Flatpickr 但不觸發 onChange 的副作用
        if (input === elements.startTime && startTimePicker) {
            startTimePicker.set('minTime', minTime);
            startTimePicker.setDate(input.value, false, 'H:i');
        } else if (input === elements.endTime && endTimePicker) {
            endTimePicker.set('minTime', minTime);
            endTimePicker.setDate(input.value, false, 'H:i');
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

// 初始化 Flatpickr 時間選取器（24小時制、5 分鐘刻度、06:00–22:00）
function initializeTimePickers() {
    if (typeof flatpickr !== 'function') return;
    const commonOptions = {
        enableTime: true,
        noCalendar: true,
        time_24hr: true,
        minuteIncrement: 5,
        wheelInput: false,
        allowInput: true,
        dateFormat: 'H:i',
        locale: (window.flatpickr && window.flatpickr.l10ns && window.flatpickr.l10ns.zh_tw) ? window.flatpickr.l10ns.zh_tw : undefined
    };
    // 讓浮層掛在各自欄位的外層 .form-group，避免跑到底部
    const startAppendTo = elements.startTime && elements.startTime.parentElement ? elements.startTime.parentElement : document.body;
    if (startAppendTo && !startAppendTo.style.position) {
        startAppendTo.style.position = 'relative';
        startAppendTo.style.overflow = 'visible';
    }
    startTimePicker = flatpickr('#startTime', {
        ...commonOptions,
        minTime: getMinTimeForSelectedDate(),
        maxTime: BUSINESS_MAX_TIME,
        defaultHour: 7,
        defaultMinute: 0,
        appendTo: startAppendTo,
        static: true,
        positionElement: elements.startTime,
        onReady: (selectedDates, dateStr, instance) => {
            ensureScrollableTimePanel(instance, 'startTime');
        },
        onOpen: () => {
            startTimePicker.set('minTime', getMinTimeForSelectedDate());
            setTimeout(() => tryFocusTimeUnitInInstance(startTimePicker, 'minute'), 0);
            refreshScrollableTimePanel(startTimePicker, 'startTime');
        },
        onChange: validateForm
    });
    const endAppendTo = elements.endTime && elements.endTime.parentElement ? elements.endTime.parentElement : document.body;
    if (endAppendTo && !endAppendTo.style.position) {
        endAppendTo.style.position = 'relative';
        endAppendTo.style.overflow = 'visible';
    }
    endTimePicker = flatpickr('#endTime', {
        ...commonOptions,
        minTime: getMinTimeForSelectedDate(),
        maxTime: BUSINESS_MAX_TIME,
        defaultHour: 8,
        defaultMinute: 0,
        appendTo: endAppendTo,
        static: true,
        positionElement: elements.endTime,
        onReady: (selectedDates, dateStr, instance) => {
            ensureScrollableTimePanel(instance, 'endTime');
        },
        onOpen: () => {
            endTimePicker.set('minTime', getMinTimeForSelectedDate());
            setTimeout(() => tryFocusTimeUnitInInstance(endTimePicker, 'minute'), 0);
            refreshScrollableTimePanel(endTimePicker, 'endTime');
        },
        onChange: validateForm
    });

    // 聚焦自動開啟時間面板
    elements.startTime.addEventListener('focus', () => startTimePicker && startTimePicker.open());
    elements.endTime.addEventListener('focus', () => endTimePicker && endTimePicker.open());

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


// 直接用 instance 內的 calendarContainer 定位聚焦，避免 appendTo 到 body 後查找不到
function tryFocusTimeUnitInInstance(instance, unit) {
    const container = instance && instance.calendarContainer;
    if (!container) return;
    const selector = unit === 'minute' ? '.numInputWrapper:nth-of-type(2) input' : '.numInputWrapper:nth-of-type(1) input';
    const target = container.querySelector(selector);
    if (target) target.focus();
}


// 新增：可滾動的小時/分鐘欄，支援滑鼠滾輪與點選
function ensureScrollableTimePanel(instance, inputId) {
    const calendar = instance && instance.calendarContainer;
    if (!calendar) return;
    let panel = calendar.querySelector('.scroll-time-panel');
    if (!panel) {
        panel = document.createElement('div');
        panel.className = 'scroll-time-panel';
        panel.style.display = 'grid';
        panel.style.gridTemplateColumns = '1fr 1fr';
        panel.style.gap = '8px';
        panel.style.marginTop = '8px';
        panel.style.position = 'relative';
        panel.style.width = '100%';
        panel.style.boxSizing = 'border-box';
        panel.style.zIndex = '10';

        const hourCol = document.createElement('div');
        const minuteCol = document.createElement('div');
        hourCol.className = 'scroll-col hours';
        minuteCol.className = 'scroll-col minutes';
        [hourCol, minuteCol].forEach(col => {
            col.style.maxHeight = '160px';
            col.style.overflowY = 'auto';
            col.style.border = '1px solid #e2e8f0';
            col.style.borderRadius = '6px';
        });

        panel.appendChild(hourCol);
        panel.appendChild(minuteCol);

        // 優先放在 Flatpickr 的時間區塊之後，避免定位錯亂
        const timeContainer = calendar.querySelector('.flatpickr-time') || calendar;
        if (timeContainer && timeContainer.parentNode) {
            if (timeContainer.nextSibling) {
                timeContainer.parentNode.insertBefore(panel, timeContainer.nextSibling);
            } else {
                timeContainer.parentNode.appendChild(panel);
            }
        } else {
            calendar.appendChild(panel);
        }

        // 確保容器可顯示溢出內容
        calendar.style.overflow = 'visible';

        // 點擊/滾動面板不關閉或影響日曆布局
        panel.addEventListener('mousedown', (e) => e.stopPropagation());
        panel.addEventListener('wheel', (e) => e.stopPropagation(), { passive: false });

        // 綁定滾輪：同步選取器時間
        const sync = () => syncFromScrollPanel(instance, inputId);
        hourCol.addEventListener('click', sync);
        minuteCol.addEventListener('click', sync);
        hourCol.addEventListener('wheel', (e) => { e.preventDefault(); hourCol.scrollTop += e.deltaY; }, { passive: false });
        minuteCol.addEventListener('wheel', (e) => { e.preventDefault(); minuteCol.scrollTop += e.deltaY; }, { passive: false });
    }
    buildScrollableTimeItems(panel, instance, inputId);
}

function buildScrollableTimeItems(panel, instance, inputId) {
    const hourCol = panel.querySelector('.scroll-col.hours');
    const minuteCol = panel.querySelector('.scroll-col.minutes');
    hourCol.innerHTML = '';
    minuteCol.innerHTML = '';
    const minTime = getMinTimeForSelectedDate();

    // 小時：06–22
    for (let h = 6; h <= 22; h++) {
        const btn = document.createElement('button');
        btn.type = 'button';
        btn.textContent = String(h).padStart(2, '0');
        styleScrollBtn(btn);
        // 若今天，且該小時整點仍小於最小允許時間，標為 disabled
        const disabled = elements.eventDate.value && compareTime(`${btn.textContent}:00`, minTime) < 0;
        if (disabled) markDisabled(btn);
        btn.addEventListener('click', () => {
            const current = getCurrentPanelTime(instance);
            const minutes = current ? current.split(':')[1] : '00';
            const val = `${btn.textContent}:${minutes}`;
            setPickerTime(instance, inputId, val);
        });
        hourCol.appendChild(btn);
    }

    // 分鐘：顯示整個 0–55 每 5 分鐘
    for (let m = 0; m < 60; m += 5) {
        const mm = String(m).padStart(2, '0');
        const btn = document.createElement('button');
        btn.type = 'button';
        btn.textContent = mm;
        styleScrollBtn(btn);
        const currentHour = getCurrentPanelTime(instance)?.split(':')[0] || '06';
        const disabled = elements.eventDate.value && compareTime(`${currentHour}:${mm}`, minTime) < 0;
        if (disabled) markDisabled(btn);
        btn.addEventListener('click', () => {
            const hour = getCurrentPanelTime(instance)?.split(':')[0] || '06';
            const val = `${hour}:${mm}`;
            setPickerTime(instance, inputId, val);
        });
        minuteCol.appendChild(btn);
    }

    // 初次同步當前時間
    highlightCurrent(panel, instance);
}

function styleScrollBtn(btn) {
    btn.className = 'scroll-time-btn';
    btn.style.display = 'block';
    btn.style.width = '100%';
    btn.style.textAlign = 'center';
    btn.style.padding = '6px 0';
    btn.style.border = 'none';
    btn.style.borderBottom = '1px solid #e2e8f0';
    btn.style.background = '#fff';
    btn.style.cursor = 'pointer';
    btn.style.fontSize = '14px';
}

function markDisabled(btn) {
    btn.disabled = true;
    btn.style.opacity = '0.5';
    btn.style.cursor = 'not-allowed';
}

function getCurrentPanelTime(instance) {
    const v = instance && instance.input && instance.input.value;
    return v && /^\d{2}:\d{2}$/.test(v) ? v : null;
}

function setPickerTime(instance, inputId, val) {
    // 夾在允許範圍（含今日動態最小值）
    const min = getMinTimeForSelectedDate();
    const max = BUSINESS_MAX_TIME;
    const adjusted = compareTime(val, min) < 0 ? min : (compareTime(val, max) > 0 ? max : val);
    const input = inputId === 'startTime' ? elements.startTime : elements.endTime;
    input.value = adjusted;
    instance.setDate(adjusted, true, 'H:i');
    validateForm();
}

function syncFromScrollPanel(instance, inputId) {
    // 點擊後已於 setPickerTime 處理
    highlightCurrent(instance.calendarContainer.querySelector('.scroll-time-panel'), instance);
}

function highlightCurrent(panel, instance) {
    if (!panel) return;
    const current = getCurrentPanelTime(instance);
    const [hh, mm] = (current || '').split(':');
    panel.querySelectorAll('.scroll-time-btn').forEach(btn => {
        btn.style.background = '#fff';
        btn.style.color = '#1e293b';
        if (btn.textContent === hh || btn.textContent === mm) {
            btn.style.background = '#2563eb';
            btn.style.color = '#fff';
        }
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

// 檢查會議室衝突
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
        // 取得 iCal（加上 corsproxy.io 代理）
        const proxy = 'https://corsproxy.io/?';
        const url = proxy + encodeURIComponent(icalUrl);
        const response = await fetch(url, {
            method: 'GET',
            headers: { 'Accept': 'text/calendar,text/plain,*/*' }
        });
        if (!response.ok) throw new Error(`無法獲取會議室日曆數據: ${response.status} ${response.statusText}`);
        const icalData = await response.text();
        const conflicts = parseICalForConflicts(icalData, startDateTime, endDateTime);
        if (conflicts.length > 0) {
            const conflictInfo = conflicts.map(c => {
                const startStr = c.start ? c.start.toLocaleString('zh-TW') : '未知時間';
                const endStr = c.end ? c.end.toLocaleString('zh-TW') : '未知時間';
                return `${c.summary || '未命名事件'} (${startStr} - ${endStr})`;
            }).join(', ');
            showNotification(`檢測到時間衝突：${conflictInfo}`, 'error');
            return true;
        }
        showNotification('會議室時間可用，無衝突', 'success');
        return false;
    } catch (error) {
        console.error('檢查會議室衝突時發生錯誤:', error);
        showNotification('無法自動檢查衝突，請手動確認會議室可用性', 'error');
        return false;
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