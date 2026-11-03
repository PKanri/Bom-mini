/* Date presentation only: application values remain YYYY-MM-DD. */
(function () {
  'use strict';
  const nativeValue = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value');
  function displayDate(value) {
    return String(value || '').replace(/^(\d{4})-(\d{2})-(\d{2})$/, '$3/$2/$1');
  }
  function parseDate(value) {
    const m = String(value).trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (!m) return '';
    const iso = `${m[3]}-${m[2].padStart(2,'0')}-${m[1].padStart(2,'0')}`;
    const date = new Date(iso + 'T00:00:00Z');
    return Number.isFinite(date.getTime()) && date.toISOString().slice(0,10) === iso ? iso : '';
  }
  window.displayDate = displayDate;
  const style = document.createElement('style');
  style.textContent = `.date-display-control{display:inline-flex;position:relative;align-items:center;max-width:100%;vertical-align:middle}.date-display-text{width:100%;min-width:0;padding-right:38px!important}.date-display-calendar{position:absolute;right:8px;top:50%;transform:translateY(-50%);width:25px;height:28px;cursor:pointer}.date-display-calendar svg{width:20px;height:20px;position:absolute;top:4px;left:2px;pointer-events:none}.date-display-calendar input[type=date]{position:absolute!important;inset:0!important;width:100%!important;height:100%!important;min-width:0!important;margin:0!important;padding:0!important;opacity:0;cursor:pointer}.date-display-calendar input::-webkit-calendar-picker-indicator{position:absolute;inset:0;width:100%;height:100%;margin:0;padding:0;cursor:pointer}.date-display-control:has(input:disabled){opacity:.6}`;
  document.head.appendChild(style);
  function enhance(input) {
    if (input.dataset.dateDisplay) return;
    input.dataset.dateDisplay = 'true';
    const wrapper = document.createElement('span');
    wrapper.className = 'date-display-control';
    const text = document.createElement('input');
    text.type = 'text';
    text.className = input.className + ' date-display-text';
    text.id = input.id ? input.id + '-display' : '';
    text.placeholder = 'วัน/เดือน/ปี';
    text.setAttribute('aria-label', 'วันที่ วัน/เดือน/ปี');
    text.autocomplete = 'off';
    const calendar = document.createElement('span');
    calendar.className = 'date-display-calendar';
    calendar.innerHTML = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" aria-hidden="true"><rect x="3" y="5" width="18" height="16" rx="2"/><path d="M7 2v6M17 2v6M3 11h18"/></svg>';
    input.before(wrapper);
    wrapper.append(text, calendar);
    calendar.append(input);
    input.setAttribute('aria-label', 'เลือกวันที่จากปฏิทิน');
    function sync() {
      text.value = displayDate(nativeValue.get.call(input));
      text.disabled = input.disabled;
      text.required = input.required;
      text.setCustomValidity('');
    }
    Object.defineProperty(input, 'value', {
      configurable: true,
      get() { return nativeValue.get.call(this); },
      set(value) { nativeValue.set.call(this, value); sync(); }
    });
    input.addEventListener('input', sync);
    input.addEventListener('change', sync);
    text.addEventListener('input', () => text.setCustomValidity(''));
    text.addEventListener('change', () => {
      const iso = parseDate(text.value);
      if (text.value.trim() && (!iso || (input.min && iso < input.min) || (input.max && iso > input.max))) {
        text.setCustomValidity('กรุณากรอกวันที่ที่ถูกต้องเป็น วัน/เดือน/ปี เช่น 03/11/2026');
        text.reportValidity();
        return;
      }
      input.value = iso;
      input.dispatchEvent(new Event('input', {bubbles:true}));
      input.dispatchEvent(new Event('change', {bubbles:true}));
    });
    new MutationObserver(sync).observe(input, {attributes:true, attributeFilter:['disabled','required','value','min','max']});
    sync();
  }
  // Format standalone ISO dates, preserving machine values and filenames.
  function formatText(node) {
    if (!node.parentElement || node.parentElement.closest('script,style,textarea,code,pre,[contenteditable]')) return;
    const original = node.nodeValue;
    const formatted = original.replace(/(^|[\s·(>])([12]\d{3}-\d{2}-\d{2})(?=$|[\s·)<,:])/g, (_, prefix, date) => prefix + displayDate(date));
    if (original !== formatted) {
      const option = node.parentElement;
      if (option.tagName === 'OPTION' && !option.hasAttribute('value')) option.value = option.textContent;
      node.nodeValue = formatted;
    }
  }
  function scan(root) {
    if (root.nodeType === Node.TEXT_NODE) return formatText(root);
    if (root.nodeType !== Node.ELEMENT_NODE) return;
    if (root.matches('input[type=date]')) enhance(root);
    root.querySelectorAll('input[type=date]').forEach(enhance);
    const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
    let node;
    while ((node = walker.nextNode())) formatText(node);
  }
  scan(document.body);
  new MutationObserver(records => {
    for (const record of records) {
      if (record.type === 'characterData') formatText(record.target);
      else record.addedNodes.forEach(scan);
    }
  }).observe(document.body, {subtree:true, childList:true, characterData:true});
})();
