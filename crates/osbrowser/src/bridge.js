// The native process owns guest code; this page owns the real DOM and pixels.
(() => {
  const documents = new Map();
  const nodes = new Map();
  const nodeIds = new WeakMap();
  const trackedDocuments = new WeakSet();
  const documentIds = new WeakMap();
  let connected = true;
  let nextNode = 1;
  let activeDocument = 0;
  const pendingEvents = [];
  let eventSending = false;
  let eventScheduled = false;
  function scheduleEvents() {
    if (!connected || eventSending || eventScheduled || !pendingEvents.length) return;
    eventScheduled = true;
    queueMicrotask(() => {
      eventScheduled = false;
      if (!connected || eventSending || !pendingEvents.length) return;
      eventSending = true;
      const batch = pendingEvents.splice(0, 128);
      fetch('event', {
        method: 'POST', headers: { 'content-type': 'application/json' },
        body: JSON.stringify(batch),
      }).catch(() => {}).finally(() => {
        eventSending = false;
        scheduleEvents();
      });
    });
  }
  // Send every queued event and wait for it to land, so a reply never overtakes
  // the events the command that produced it dispatched.
  async function drainEvents() {
    while (eventSending) await new Promise(resolve => setTimeout(resolve, 0));
    while (connected && pendingEvents.length) {
      eventSending = true;
      const batch = pendingEvents.splice(0, 512);
      try {
        await fetch('event', { method: 'POST', headers: { 'content-type': 'application/json' },
          body: JSON.stringify(batch) });
      } catch (_) {
      } finally {
        eventSending = false;
      }
    }
  }

  const idOf = node => {
    if (!node) return 0;
    let id = nodeIds.get(node);
    if (!id) {
      id = nextNode++;
      nodeIds.set(node, id);
    }
    nodes.set(id, node);
    return id;
  };
  const docOf = id => documents.get(id) || document;

  const reportViewport = docId => {
    pendingEvents.push({ document: docId, node: 0, kind: 'resize',
      fields: { width: innerWidth, height: innerHeight } });
    scheduleEvents();
  };
  const reportScroll = docId => {
    pendingEvents.push({ document: docId, node: 0, kind: 'scroll',
      fields: { scroll_x: scrollX, scroll_y: scrollY } });
    scheduleEvents();
  };

  const trackEvents = (doc, docId) => {
    documentIds.set(doc, docId);
    if (doc === document) reportViewport(docId);
    if (trackedDocuments.has(doc)) return;
    trackedDocuments.add(doc);
    if (doc === document) {
      window.addEventListener('resize', () => connected && reportViewport(documentIds.get(doc)));
      window.addEventListener('scroll', () => connected && reportScroll(documentIds.get(doc)), { passive: true });
    }
    for (const kind of [
      'click', 'input', 'change', 'submit', 'focus', 'blur', 'keydown', 'keyup',
      'mousedown', 'mouseup', 'mousemove', 'wheel', 'close', 'load'
    ]) {
      doc.addEventListener(kind, event => {
        if (!connected) return;
        const target = event.target === doc ? doc.body : event.target;
        const fields = {
          key: event.key || '', code: event.code || '', key_code: event.keyCode || 0,
          client_x: event.clientX || 0, client_y: event.clientY || 0,
          button: event.button || 0, buttons: event.buttons || 0,
          delta_y: event.deltaY || 0, ctrl_key: Boolean(event.ctrlKey),
          shift_key: Boolean(event.shiftKey), alt_key: Boolean(event.altKey),
          meta_key: Boolean(event.metaKey), repeat: Boolean(event.repeat)
        };
        fields.form = formOf(target);
        pendingEvents.push({
            document: documentIds.get(doc), node: idOf(target), kind, fields,
            path: event.bubbles
              ? event.composedPath().filter(n => n === doc || n.nodeType === 1)
                  .map(n => n === doc ? 0 : idOf(n))
              : [idOf(target)]
          });
        scheduleEvents();
      }, true);
    }
  };

  const disconnect = () => {
    connected = false;
    pendingEvents.length = 0;
    console.warn('osbrowser: the bridge disconnected; reload after restarting the app.');
  };

  // ── webcore's Document API ──────────────────────────────────────────────
  //
  // One entry per `osbrowser::types::Document` method, named after it. Node
  // ids follow webcore: `0` is "no node", never the body.
  const N = id => id ? nodes.get(id) || null : null;
  const id0 = node => node ? idOf(node) : 0;
  const ids = list => Array.from(list || [], idOf);
  const rootId = doc => id0(doc.documentElement);
  const nodeIdIn = (doc, node) => node === doc ? rootId(doc) : id0(node);
  const select = (scope, selector) => { try { return scope.querySelector(selector); } catch (_) { return null; } };
  const selectAll = (scope, selector) => { try { return scope.querySelectorAll(selector); } catch (_) { return []; } };
  const rect = r => ({ x: r.x, y: r.y, w: r.width, h: r.height });
  const listenerFns = new Map();
  let pressed = null;
  let hovered = null;

  const eventFields = (doc, event) => ({
    phase: event.eventPhase,
    key: event.key || '', code: event.code || '', key_code: event.keyCode || 0,
    client_x: event.clientX || 0, client_y: event.clientY || 0,
    button: event.button || 0, buttons: event.buttons || 0,
    delta_x: event.deltaX || 0, delta_y: event.deltaY || 0,
    ctrl_key: Boolean(event.ctrlKey), shift_key: Boolean(event.shiftKey),
    alt_key: Boolean(event.altKey), meta_key: Boolean(event.metaKey),
    repeat: Boolean(event.repeat), is_trusted: event.isTrusted,
    time_stamp: event.timeStamp,
    composed_path: event.composedPath().filter(n => n === doc || n.nodeType)
      .map(n => nodeIdIn(doc, n)),
  });

  const formOf = node => {
    if (!node || !node.localName) return undefined;
    const tag = node.localName;
    if (!['input', 'select', 'textarea', 'button', 'form'].includes(tag)) return undefined;
    return { tag, id: node.id || '', name: node.getAttribute('name') || '',
      type: (node.type || '').toLowerCase(), value: node.value ?? '',
      checked: Boolean(node.checked), action: tag === 'form' ? node.action : '' };
  };

  const hitAt = (doc, x, y) => {
    const cx = x - (doc.defaultView?.scrollX || 0), cy = y - (doc.defaultView?.scrollY || 0);
    return { cx, cy, target: doc.elementFromPoint(cx, cy) || doc.body };
  };

  const pointerInit = (cx, cy, button) => ({ bubbles: true, cancelable: true, composed: true,
    clientX: cx, clientY: cy, button, buttons: button === 0 ? 1 : button === 2 ? 2 : 4 });

  const insertText = (target, text) => {
    if (!('setRangeText' in target)) return false;
    const start = target.selectionStart ?? target.value.length;
    const end = target.selectionEnd ?? start;
    if (text === null) {
      if (start === end && start > 0) target.setRangeText('', start - 1, end, 'end');
      else target.setRangeText('', start, end, 'end');
    } else {
      target.setRangeText(text, start, end, 'end');
    }
    target.dispatchEvent(new InputEvent('input', { bubbles: true, composed: true }));
    return true;
  };

  const methods = {
    load_html(doc, [html]) {
      const parsed = new DOMParser().parseFromString(html, 'text/html');
      doc.replaceChild(doc.importNode(parsed.documentElement, true), doc.documentElement);
      return rootId(doc);
    },
    title: doc => doc.title,
    set_title(doc, [title]) { doc.title = title; },
    body: doc => id0(doc.body),
    head: doc => id0(doc.head),
    document_element: doc => rootId(doc),
    facts(doc, [id]) {
      const n = N(id);
      if (!n) return null;
      return { tag: n.nodeType === 1 ? n.tagName : null, local: n.localName || '',
        name: n.nodeName, namespace: n.namespaceURI ?? null, prefix: n.prefix ?? null,
        node_type: n.nodeType };
    },

    create_element: (doc, [tag]) => idOf(doc.createElement(tag)),
    create_text_node: (doc, [text]) => idOf(doc.createTextNode(text)),
    create_comment: (doc, [text]) => idOf(doc.createComment(text)),
    create_cdata_section: (doc, [data]) => idOf(doc.createCDATASection(data)),
    create_element_ns: (doc, [ns, name]) => idOf(doc.createElementNS(ns || null, name)),
    create_processing_instruction: (doc, [target, data]) => idOf(doc.createProcessingInstruction(target, data)),
    create_document_fragment: doc => idOf(doc.createDocumentFragment()),
    clone_node: (doc, [id, deep]) => idOf(N(id).cloneNode(deep)),

    get_element_by_id: (doc, [id]) => id0(doc.getElementById(id)),
    get_elements_by_tag_name: (doc, [tag]) => ids(doc.getElementsByTagName(tag)),
    get_elements_by_class_name: (doc, [names]) => ids(doc.getElementsByClassName(names)),
    query_selector: (doc, [selector]) => id0(select(doc, selector)),
    query_selector_all: (doc, [selector]) => ids(selectAll(doc, selector)),
    query_selector_within: (doc, [root, selector]) => id0(N(root) && select(N(root), selector)),
    query_selector_all_within: (doc, [root, selector]) => N(root) ? ids(selectAll(N(root), selector)) : [],
    matches(doc, [id, selectors]) { try { return Boolean(N(id)?.matches(selectors)); } catch (_) { return false; } },
    closest(doc, [id, selectors]) { try { return id0(N(id)?.closest(selectors)); } catch (_) { return 0; } },
    element_from_point: (doc, [x, y]) => id0(doc.elementFromPoint(x, y)),

    append_child(doc, [parent, child]) { N(parent)?.appendChild(N(child)); },
    insert_before(doc, [parent, child, reference]) {
      const p = N(parent), r = N(reference);
      if (p && r?.parentNode === p) p.insertBefore(N(child), r);
    },
    remove_child(doc, [child]) { N(child)?.remove(); },
    replace_child(doc, [parent, fresh, old]) {
      const p = N(parent), o = N(old);
      if (!p || o?.parentNode !== p) return false;
      p.replaceChild(N(fresh), o);
      return true;
    },
    append(doc, [parent, list]) { N(parent)?.append(...list.map(N)); },
    prepend(doc, [parent, list]) { N(parent)?.prepend(...list.map(N)); },
    before(doc, [id, list]) { N(id)?.before(...list.map(N)); },
    after(doc, [id, list]) { N(id)?.after(...list.map(N)); },
    replace_with(doc, [id, list]) { N(id)?.replaceWith(...list.map(N)); },
    insert_adjacent_element: (doc, [id, position, other]) => id0(N(id)?.insertAdjacentElement(position, N(other))),
    insert_adjacent_text(doc, [id, position, data]) { N(id)?.insertAdjacentText(position, data); },
    normalize(doc, [id]) { N(id)?.normalize(); },
    parent_node: (doc, [id]) => nodeIdIn(doc, N(id)?.parentNode),
    parent_element: (doc, [id]) => id0(N(id)?.parentElement),
    child_nodes: (doc, [id]) => ids(N(id)?.childNodes),
    children: (doc, [id]) => ids(N(id)?.children),
    first_child: (doc, [id]) => id0(N(id)?.firstChild),
    last_child: (doc, [id]) => id0(N(id)?.lastChild),
    next_sibling: (doc, [id]) => id0(N(id)?.nextSibling),
    previous_sibling: (doc, [id]) => id0(N(id)?.previousSibling),
    first_element_child: (doc, [id]) => id0(N(id)?.firstElementChild),
    last_element_child: (doc, [id]) => id0(N(id)?.lastElementChild),
    next_element_sibling: (doc, [id]) => id0(N(id)?.nextElementSibling),
    previous_element_sibling: (doc, [id]) => id0(N(id)?.previousElementSibling),
    child_element_count: (doc, [id]) => N(id)?.childElementCount || 0,
    has_child_nodes: (doc, [id]) => Boolean(N(id)?.hasChildNodes()),
    contains: (doc, [id, other]) => Boolean(N(id)?.contains(N(other))),
    is_connected: (doc, [id]) => Boolean(N(id)?.isConnected),
    is_equal_node: (doc, [a, b]) => Boolean(N(a)?.isEqualNode(N(b))),
    compare_document_position: (doc, [a, b]) => N(a) && N(b) ? N(a).compareDocumentPosition(N(b)) : 0,

    node_value: (doc, [id]) => N(id)?.nodeValue ?? null,
    text_content: (doc, [id]) => N(id)?.textContent ?? '',
    set_text_content(doc, [id, text]) { const n = N(id); if (n) n.textContent = text; },
    text_data: (doc, [id]) => N(id)?.data ?? '',
    set_text_data(doc, [id, data]) { const n = N(id); if (n && 'data' in n) n.data = data; },
    inner_html: (doc, [id]) => N(id)?.innerHTML ?? '',
    outer_html: (doc, [id]) => N(id)?.outerHTML ?? '',
    inner_text: (doc, [id]) => N(id)?.innerText ?? N(id)?.textContent ?? '',

    get_attribute: (doc, [id, key]) => N(id)?.getAttribute?.(key) ?? null,
    set_attribute(doc, [id, key, value]) { N(id)?.setAttribute(key, value); },
    remove_attribute(doc, [id, key]) { N(id)?.removeAttribute(key); },
    has_attribute: (doc, [id, name]) => Boolean(N(id)?.hasAttribute?.(name)),
    has_attributes: (doc, [id]) => Boolean(N(id)?.hasAttributes?.()),
    toggle_attribute: (doc, [id, name]) => Boolean(N(id)?.toggleAttribute(name)),
    get_attribute_names: (doc, [id]) => N(id)?.getAttributeNames?.() || [],
    get_attribute_ns: (doc, [id, ns, local]) => N(id)?.getAttributeNS?.(ns || null, local) ?? null,
    set_attribute_ns(doc, [id, ns, name, value]) { N(id)?.setAttributeNS(ns || null, name, value); },
    has_attribute_ns: (doc, [id, ns, local]) => Boolean(N(id)?.hasAttributeNS?.(ns || null, local)),
    remove_attribute_ns(doc, [id, ns, local]) { N(id)?.removeAttributeNS(ns || null, local); },
    id: (doc, [id]) => N(id)?.id ?? '',
    set_id(doc, [id, value]) { const n = N(id); if (n) n.id = value; },
    class_name: (doc, [id]) => N(id)?.getAttribute?.('class') ?? '',
    set_class_name(doc, [id, value]) { N(id)?.setAttribute('class', value); },
    class_list_add(doc, [id, name]) { N(id)?.classList.add(name); },
    class_list_remove(doc, [id, name]) { N(id)?.classList.remove(name); },
    class_list_contains: (doc, [id, name]) => Boolean(N(id)?.classList?.contains(name)),
    class_list_toggle: (doc, [id, name]) => Boolean(N(id)?.classList.toggle(name)),
    dataset: (doc, [id]) => Object.entries(N(id)?.dataset || {}).sort((a, b) => a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : 0),
    dataset_get: (doc, [id, key]) => N(id)?.dataset?.[key] ?? null,
    set_dataset(doc, [id, key, value]) { const n = N(id); if (n?.dataset) n.dataset[key] = value; },
    remove_dataset(doc, [id, key]) { const n = N(id); if (n?.dataset) delete n.dataset[key]; },
    tab_index: (doc, [id]) => N(id)?.tabIndex ?? -1,
    set_tab_index(doc, [id, index]) { const n = N(id); if (n) n.tabIndex = index; },

    get_style_property(doc, [id, prop]) {
      const n = N(id);
      if (!n?.hasAttribute?.('style')) return null;
      const value = n.style.getPropertyValue(prop);
      return value === '' ? null : value;
    },
    set_style_property(doc, [id, prop, value]) { N(id)?.style.setProperty(prop, value); },
    remove_style_property(doc, [id, prop]) {
      const value = N(id)?.style?.removeProperty(prop);
      return value ? value : null;
    },
    style_property_len: (doc, [id]) => N(id)?.style?.length || 0,
    style_property_item: (doc, [id, index]) => N(id)?.style?.item(index) || null,
    computed_style_property(doc, [id, prop]) {
      const n = N(id);
      return n?.nodeType === 1 ? (doc.defaultView || window).getComputedStyle(n).getPropertyValue(prop) : '';
    },

    get_bounding_client_rect(doc, [id]) {
      const n = N(id);
      return n?.getBoundingClientRect ? rect(n.getBoundingClientRect()) : null;
    },
    get_client_rects: (doc, [id]) => Array.from(N(id)?.getClientRects?.() || [], rect),
    offset_width: (doc, [id]) => N(id)?.offsetWidth || 0,
    offset_height: (doc, [id]) => N(id)?.offsetHeight || 0,
    offset_top: (doc, [id]) => N(id)?.offsetTop || 0,
    offset_left: (doc, [id]) => N(id)?.offsetLeft || 0,
    offset_parent: (doc, [id]) => id0(N(id)?.offsetParent),
    scroll_into_view(doc, [id]) { N(id)?.scrollIntoView(); },

    value: (doc, [id]) => N(id)?.value ?? '',
    set_value(doc, [id, value]) { const n = N(id); if (n) n.value = value; },
    checked: (doc, [id]) => Boolean(N(id)?.checked),
    set_checked(doc, [id, checked]) { const n = N(id); if (n) n.checked = checked; },
    add_item(doc, [id, text]) { N(id)?.add(new Option(text, text)); },
    remove_item(doc, [id, index]) { N(id)?.remove(index); },
    clear_items(doc, [id]) { const n = N(id); if (n?.options) n.options.length = 0; },
    item_text: (doc, [id, index]) => N(id)?.options?.[index]?.text ?? '',
    set_item_text(doc, [id, index, text]) { const o = N(id)?.options?.[index]; if (o) o.text = text; },
    selected_index: (doc, [id]) => N(id)?.selectedIndex ?? -1,
    set_selected_index(doc, [id, index]) { const n = N(id); if (n) n.selectedIndex = index; },
    focus(doc, [id]) { N(id)?.focus?.(); },
    click(doc, [id]) { N(id)?.click?.(); return true; },
    show_dialog(doc, [id, modal]) { const n = N(id); if (n && !n.open) modal ? n.showModal() : n.show(); },
    close_dialog(doc, [id]) { N(id)?.close?.(); },
    dialog_open: (doc, [id]) => Boolean(N(id)?.open),

    add_event_listener(doc, [id, type, listener, capture, once, passive], docId) {
      const target = N(id);
      if (!target) return;
      const fn = event => {
        if (!connected) return;
        pendingEvents.push({ document: docId, node: nodeIdIn(doc, event.target), kind: event.type,
          listener, fields: eventFields(doc, event) });
        scheduleEvents();
      };
      listenerFns.set(listener, { target, type, fn, capture });
      target.addEventListener(type, fn, { capture, once, passive });
    },
    remove_event_listener(doc, [listener]) {
      const entry = listenerFns.get(listener);
      if (!entry) return;
      entry.target.removeEventListener(entry.type, entry.fn, { capture: entry.capture });
      listenerFns.delete(listener);
    },

    process_mouse_event(doc, [type, x, y, button]) {
      const { cx, cy, target } = hitAt(doc, x, y);
      const init = pointerInit(cx, cy, button);
      const Kind = type.startsWith('pointer') ? PointerEvent : MouseEvent;
      const proceed = target.dispatchEvent(new Kind(type, init));
      if (type === 'mousedown') {
        pressed = target;
        if (proceed && target.focus) target.focus();
      } else if (type === 'mouseup') {
        if (pressed && (pressed === target || pressed.contains(target)) && button === 0) target.click();
        pressed = null;
      }
      return true;
    },
    dispatch_over_out(doc, [x, y]) {
      const { cx, cy, target } = hitAt(doc, x, y);
      if (target === hovered) return false;
      const init = { ...pointerInit(cx, cy, 0), buttons: 0 };
      if (hovered) {
        hovered.dispatchEvent(new MouseEvent('mouseout', { ...init, relatedTarget: target }));
        hovered.dispatchEvent(new MouseEvent('mouseleave', { ...init, bubbles: false, relatedTarget: target }));
      }
      target.dispatchEvent(new MouseEvent('mouseover', { ...init, relatedTarget: hovered }));
      target.dispatchEvent(new MouseEvent('mouseenter', { ...init, bubbles: false, relatedTarget: hovered }));
      hovered = target;
      return true;
    },
    process_keyboard_event(doc, [e]) {
      const target = doc.activeElement || doc.body;
      const proceed = target.dispatchEvent(new KeyboardEvent(e.kind, { bubbles: true, cancelable: true,
        composed: true, key: e.key, code: e.code, repeat: e.repeat, ctrlKey: e.ctrl_key,
        shiftKey: e.shift_key, altKey: e.alt_key, metaKey: e.meta_key }));
      if (!proceed || e.kind !== 'keydown' || e.ctrl_key || e.meta_key) return proceed;
      if (e.key === 'Backspace') insertText(target, null);
      else if ([...e.key].length === 1) insertText(target, e.key);
      return proceed;
    },
    process_wheel_event_xy(doc, [x, y, dx, dy]) {
      let el = hitAt(doc, x, y).target;
      while (el && el !== doc.body && !(el.scrollHeight > el.clientHeight || el.scrollWidth > el.clientWidth)) {
        el = el.parentElement;
      }
      const scroller = !el || el === doc.body ? (doc.scrollingElement || doc.documentElement) : el;
      const before = [scroller.scrollLeft, scroller.scrollTop];
      scroller.scrollBy(-dx, -dy);
      return scroller.scrollLeft !== before[0] || scroller.scrollTop !== before[1];
    },
    dispatch_input_event(doc, [e]) {
      const target = N(e.target) || doc.elementFromPoint(e.client_x, e.client_y) || doc.body;
      const init = { bubbles: true, cancelable: true, composed: true, clientX: e.client_x,
        clientY: e.client_y, button: e.button, ctrlKey: e.ctrl_key, shiftKey: e.shift_key,
        altKey: e.alt_key, metaKey: e.meta_key };
      const event = e.kind === 'wheel'
        ? new WheelEvent('wheel', { ...init, deltaX: e.delta_x, deltaY: e.delta_y })
        : new MouseEvent(e.kind, init);
      const proceed = target.dispatchEvent(event);
      return { dispatched: true, prevented: !proceed };
    },

    canvas_draw(doc, [id, commands]) {
      const canvas = N(id);
      if (canvas?.getContext) window.osbrowserCanvas2D.draw(canvas.getContext('2d'), canvas, commands);
    },
    canvas_query(doc, [id, question, args]) {
      const canvas = N(id);
      return canvas?.getContext ? window.osbrowserCanvas2D.query(canvas.getContext('2d'), canvas, question, args) : null;
    },
    canvas_parse_color: (doc, [css]) => window.osbrowserCanvas2D.parseColor(css),
    canvas_parse_font: (doc, [css]) => window.osbrowserCanvas2D.parseFont(css),

    window_focus() { window.focus(); },
    window_screen_size: () => [screen.width, screen.height],
    window_resize_to(doc, [width, height]) { window.resizeTo(width, height); },
    window_move_to(doc, [x, y]) { window.moveTo(x, y); },
    window_screen_position: () => [screenX, screenY],
    window_alert(doc, [message]) { alert(message); return true; },
    window_confirm: (doc, [message]) => confirm(message),
  };

  function execute(command) {
    const { document: docId, operation: op, args: a } = command;
    if (Object.hasOwn(methods, op)) return methods[op](docOf(docId), a, docId);
    if (op === 'canvas_image') {
      const canvas = nodes.get(a.node);
      if (canvas?.getContext) window.osbrowserCanvas2D.image(canvas.getContext('2d'), a.command, a.pixels);
      return null;
    }
    if (op === 'NewDocument') {
      const doc = activeDocument === 0
        ? document
        : document.implementation.createHTMLDocument(a.title || '');
      documents.set(docId, doc);
      if (activeDocument === 0) activeDocument = docId;
      doc.title = a.title || '';
      trackEvents(doc, docId);
      return 'None';
    }
    if (op === 'Reset') {
      for (const { target, type, fn, capture } of listenerFns.values()) {
        target.removeEventListener(type, fn, { capture });
      }
      listenerFns.clear(); pressed = null; hovered = null;
      documents.clear(); nodes.clear(); activeDocument = 0;
      document.body.replaceChildren();
      return 'None';
    }
    throw new Error(`unsupported operation: ${op}`);
  }

  async function pump() {
    let failures = 0;
    for (;;) {
      try {
        const response = await fetch('next', { cache: 'no-store' });
        if (!response.ok && response.status !== 204) {
          throw new Error(`browser bridge returned ${response.status}`);
        }
        failures = 0;
        if (response.status === 204) {
          await new Promise(resolve => setTimeout(resolve, 20));
          continue;
        }
        let commands;
        if (response.headers.get('content-type') === 'application/octet-stream') {
          const buffer = await response.arrayBuffer();
          if (buffer.byteLength < 4) throw new Error('truncated binary command header');
          const length = new DataView(buffer).getUint32(0, true);
          const start = 4 + length;
          if (start > buffer.byteLength) throw new Error('truncated binary commands');
          commands = JSON.parse(new TextDecoder().decode(new Uint8Array(buffer, 4, length)));
          for (const command of commands) {
            if (!command.binary_range) continue;
            const [offset, size] = command.binary_range;
            if (!Number.isSafeInteger(offset) || !Number.isSafeInteger(size) ||
                offset < 0 || size < 0 || offset + size > buffer.byteLength - start)
              throw new Error('invalid binary pixel range');
            if (command.operation === 'canvas_image') {
              command.args.pixels = new Uint8ClampedArray(buffer, start + offset, size);
              continue;
            }
            throw new Error('binary pixels are only carried by canvas_image');
          }
        } else {
          commands = await response.json();
        }
        const replies = [];
        for (const command of commands) {
          let resultValue = 'None', error;
          try {
            resultValue = execute(command);
            if (resultValue === undefined) resultValue = null;
            if (command.assigned_node != null) {
              const originalId = typeof resultValue === 'number' ? resultValue : resultValue.Node;
              const node = nodes.get(originalId);
              if (!node) throw new Error('browser did not create a node');
              nodes.delete(originalId);
              nodes.set(command.assigned_node, node);
              nodeIds.set(node, command.assigned_node);
              if (typeof resultValue === 'number') resultValue = command.assigned_node;
            }
          } catch (cause) {
            error = String(cause);
            if (!command.id) console.error(`osbrowser ${command.operation}: ${error}`);
          }
          if (command.id) replies.push({ id: command.id, result: resultValue, error });
        }
        if (replies.length) {
          await drainEvents();
          await fetch('reply', { method: 'POST', headers: { 'content-type': 'application/json' },
            body: JSON.stringify(replies) });
        }
      } catch (_) {
        if (++failures >= 3) {
          disconnect();
          return;
        }
        await new Promise(resolve => setTimeout(resolve, 250 * failures));
      }
    }
  }
  pump();
})();
