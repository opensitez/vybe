// The native process owns guest code; this page owns the real DOM and pixels.
(() => {
  const documents = new Map();
  const nodes = new Map();
  const nodeIds = new WeakMap();
  const trackedDocuments = new WeakSet();
  const documentIds = new WeakMap();
  let nextNode = 1;
  let activeDocument = 0;
  let nextSchedule = 1;
  const timers = new Map();
  const dueTimers = [];
  const frames = new Map();
  const dueFrames = [];
  const pointer = { client_x: 0, client_y: 0, buttons: 0,
    ctrl_key: false, shift_key: false, alt_key: false, meta_key: false };

  const result = (kind, value) => value === undefined ? kind : { [kind]: value };
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
  const nodeOf = (id, doc) => id === 0 ? doc.body : nodes.get(id);
  const text = value => result('Text', String(value ?? ''));
  const bool = value => result('Bool', Boolean(value));
  const number = value => result('Number', Number(value));
  const nodeResult = value => value ? result('Node', idOf(value)) : 'Null';
  const nodesResult = value => result('Nodes', Array.from(value, idOf));
  const absent = value => value == null ? 'Null' : text(value);

  const trackEvents = (doc, docId) => {
    documentIds.set(doc, docId);
    if (trackedDocuments.has(doc)) return;
    trackedDocuments.add(doc);
    for (const kind of [
      'click', 'input', 'change', 'submit', 'focus', 'blur', 'keydown', 'keyup',
      'mousedown', 'mouseup', 'mousemove', 'wheel', 'close', 'load'
    ]) {
      doc.addEventListener(kind, event => {
        if ('clientX' in event) {
          Object.assign(pointer, { client_x: event.clientX, client_y: event.clientY,
            buttons: event.buttons, ctrl_key: event.ctrlKey,
            shift_key: event.shiftKey, alt_key: event.altKey, meta_key: event.metaKey });
        }
        const target = event.target === doc ? doc.body : event.target;
        const fields = {
          key: event.key || '', code: event.code || '', key_code: event.keyCode || 0,
          client_x: event.clientX || 0, client_y: event.clientY || 0,
          button: event.button || 0, buttons: event.buttons || 0,
          delta_y: event.deltaY || 0, ctrl_key: Boolean(event.ctrlKey),
          shift_key: Boolean(event.shiftKey), alt_key: Boolean(event.altKey),
          meta_key: Boolean(event.metaKey), repeat: Boolean(event.repeat)
        };
        fetch('event', {
          method: 'POST', headers: { 'content-type': 'application/json' },
          body: JSON.stringify({ document: documentIds.get(doc), node: idOf(target), kind, fields })
        }).catch(() => {});
      }, true);
    }
  };

  function execute(command) {
    const { document: docId, operation: op, args: a } = command;
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
      for (const entry of timers.values()) clearTimeout(entry.handle);
      for (const handle of frames.values()) cancelAnimationFrame(handle);
      timers.clear(); frames.clear(); dueTimers.length = 0; dueFrames.length = 0;
      documents.clear(); nodes.clear(); activeDocument = 0;
      document.body.replaceChildren();
      return 'None';
    }
    if (op === 'SetDocumentKind') {
      if (a === 'xml') documents.set(docId, new DOMParser().parseFromString('<root/>', 'application/xml'));
      return 'None';
    }
    if (op.startsWith('Canvas')) return window.vybeCanvas(op, a, docOf(docId), nodes);
    if (op === 'EventPointerState') return result('Pointer', pointer);
    if (op === 'ScheduleSetTimer') {
      const id = nextSchedule++;
      const delay = Math.max(0, Number(a) || 0);
      const deadline = performance.now() + delay;
      const handle = setTimeout(() => { timers.delete(id); dueTimers.push(id); }, delay);
      timers.set(id, { handle, deadline });
      return result('Id', id);
    }
    if (op === 'ScheduleClearTimer') {
      const entry = timers.get(a);
      if (entry) clearTimeout(entry.handle);
      timers.delete(a);
      return result('Bool', Boolean(entry));
    }
    if (op === 'ScheduleTakeDueTimer') return dueTimers.length ? result('Id', dueTimers.shift()) : 'Null';
    if (op === 'ScheduleRequestFrame') {
      const id = nextSchedule++;
      const handle = requestAnimationFrame(() => { frames.delete(id); dueFrames.push(id); });
      frames.set(id, handle);
      return result('Id', id);
    }
    if (op === 'ScheduleCancelFrame') {
      const handle = frames.get(a);
      if (handle !== undefined) cancelAnimationFrame(handle);
      frames.delete(a);
      return result('Bool', handle !== undefined);
    }
    if (op === 'ScheduleTakeDueFrame') return dueFrames.length ? result('Id', dueFrames.shift()) : 'Null';
    if (op === 'ScheduleTimerDelayMs') {
      if (dueTimers.length) return result('Ms', 0);
      const times = Array.from(timers.values(), entry => entry.deadline - performance.now());
      return times.length ? result('Ms', Math.max(0, Math.min(...times))) : 'Null';
    }
    if (op === 'ScheduleFrameDelayMs') return frames.size || dueFrames.length ? result('Ms', 16) : 'Null';
    if (op === 'ScheduleNow') return result('Ms', performance.now());
    if (op === 'WindowClosed') return result('Bool', !documents.has(a));
    if (op === 'WindowClose') { documents.delete(a); return 'None'; }
    if (op === 'WindowFocus') { window.focus(); return 'None'; }
    if (op === 'WindowScreen') return result('Pair', [screen.width, screen.height]);
    if (op === 'WindowInnerSize') return result('Pair', [innerWidth, innerHeight]);
    if (op === 'WindowResizeTo') { window.resizeTo(a[1], a[2]); return 'None'; }
    if (op === 'WindowViewportChanged') return 'None';
    if (op === 'WindowMoveTo') { window.moveTo(a[1], a[2]); return 'None'; }
    if (op === 'WindowScreenPosition') return result('Pair', [screenX, screenY]);
    if (op === 'WindowName') return result('Text', docOf(a).title || '');
    if (op === 'WindowAlert') { alert(a); return 'None'; }
    if (op === 'WindowConfirm') return result('Bool', confirm(a));
    const doc = docOf(docId);
    const n = id => nodeOf(id, doc);
    const first = (name, scope = doc) => {
      try { return scope.querySelector(name); } catch (_) { return null; }
    };
    const all = (name, scope = doc) => {
      try { return scope.querySelectorAll(name); } catch (_) { return []; }
    };
    switch (op) {
      case 'CreateElement': {
        const el = doc.createElement(a.tag);
        if (a.input_type) el.setAttribute('type', a.input_type);
        return nodeResult(el);
      }
      case 'CreateTextNode': return nodeResult(doc.createTextNode(a));
      case 'CreateComment': return nodeResult(doc.createComment(a));
      case 'CreateCDataSection': return nodeResult(doc.createCDATASection(a));
      case 'CreateElementNS': {
        const el = doc.createElementNS(a.namespace || null, a.qualified_name);
        if (a.input_type) el.setAttribute('type', a.input_type);
        return nodeResult(el);
      }
      case 'CreateProcessingInstruction':
        return nodeResult(doc.createProcessingInstruction(a.target, a.data));
      case 'NamespaceUri': return absent(n(a)?.namespaceURI);
      case 'Prefix': return absent(n(a)?.prefix);
      case 'LocalName': return text(n(a)?.localName || '');
      case 'CanvasSize': {
        const el = n(a);
        return result('Pair', [el?.width ?? 300, el?.height ?? 150]);
      }
      case 'GetElementById': return nodeResult(doc.getElementById(a));
      case 'ElementsByTag': return nodesResult(doc.getElementsByTagName(a));
      case 'QuerySelector': return nodeResult(first(a));
      case 'QuerySelectorAll': return nodesResult(all(a));
      case 'Title': return text(doc.title);
      case 'SetTitle': doc.title = a; return 'None';
      case 'AppendChild': n(a.parent)?.appendChild(n(a.child)); return 'None';
      case 'RemoveChild': n(a.parent)?.removeChild(n(a.child)); return 'None';
      case 'InsertBefore': {
        const parent = n(a.parent), ref = n(a.reference);
        if (!parent || ref?.parentNode !== parent) return bool(false);
        parent.insertBefore(n(a.child), ref);
        return bool(true);
      }
      case 'ReplaceChild': n(a.parent)?.replaceChild(n(a.new_child), n(a.old_child)); return 'None';
      case 'CloneNode': return nodeResult(n(a.node)?.cloneNode(a.deep));
      case 'NodeType': return number(a === 0 ? 9 : n(a)?.nodeType || 0);
      case 'NodeName': return text(a === 0 ? '#document' : n(a)?.nodeName || '');
      case 'NodeValue': return absent(a === 0 ? null : n(a)?.nodeValue);
      case 'ParentNode': return nodeResult(n(a)?.parentNode);
      case 'ChildNodes': return nodesResult(n(a)?.childNodes || []);
      case 'InnerHtml': return text(n(a)?.innerHTML || '');
      case 'SetInnerHtml': n(a.node).innerHTML = a.html; return 'None';
      case 'OuterHtml': return text(n(a)?.outerHTML || '');
      case 'SetOuterHtml': n(a.node).outerHTML = a.html; return 'None';
      case 'InsertAdjacentHtml': n(a.node)?.insertAdjacentHTML(a.position, a.html); return 'None';
      case 'CreateDocumentFragment': return nodeResult(doc.createDocumentFragment());
      case 'ImportNode': return nodeResult(doc.importNode(nodes.get(a.node), a.deep));
      case 'IsConnected': return bool(n(a)?.isConnected);
      case 'TextContent': return absent(n(a)?.textContent);
      case 'SetTextContent': n(a[0]).textContent = a[1]; return 'None';
      case 'SetAttribute': n(a[0])?.setAttribute(a[1], a[2]); return 'None';
      case 'GetAttribute': return absent(n(a[0])?.getAttribute(a[1]));
      case 'AttributeNames': return result('Texts', n(a)?.getAttributeNames() || []);
      case 'RemoveAttribute': n(a[0])?.removeAttribute(a[1]); return 'None';
      case 'SetAttributeNS':
        n(a.node)?.setAttributeNS(a.namespace || null, a.qualified_name, a.value);
        return 'None';
      case 'GetAttributeNS':
        return absent(n(a.node)?.getAttributeNS(a.namespace || null, a.local_name));
      case 'SetStyleProperty': n(a[0])?.style.setProperty(a[1], a[2]); return 'None';
      case 'GetStyleProperty': return text(n(a[0])?.style.getPropertyValue(a[1]) || '');
      case 'StyleDeclarations': {
        const style = n(a)?.style;
        return result('Properties', style ? Array.from(style, key => [key, style.getPropertyValue(key)]) : []);
      }
      case 'ComputedStyleProperty': return text(getComputedStyle(n(a[0])).getPropertyValue(a[1]));
      case 'Focus': n(a)?.focus(); return 'None';
      case 'Value': return text(n(a)?.value || '');
      case 'SetValue': n(a[0]).value = a[1]; return 'None';
      case 'Checked': return bool(n(a)?.checked);
      case 'SetChecked': n(a[0]).checked = a[1]; return 'None';
      case 'SelectedIndex': return number(n(a)?.selectedIndex ?? -1);
      case 'SetSelectedIndex': n(a[0]).selectedIndex = a[1]; return 'None';
      case 'ItemText': return text(n(a[0])?.options?.[a[1]]?.text || '');
      case 'SetItemText': n(a[0]).options[a[1]].text = a[2]; return 'None';
      case 'AddItem': n(a[0])?.add(new Option(a[1], a[1])); return 'None';
      case 'RemoveItem': n(a[0])?.remove(a[1]); return 'None';
      case 'ClearItems': n(a).replaceChildren(); return 'None';
      case 'ShowDialog': a.modal ? n(a.node)?.showModal() : n(a.node)?.show(); return 'None';
      case 'CloseDialog': n(a)?.close(); return 'None';
      case 'DialogOpen': return bool(n(a)?.open);
      case 'ShowPicker': n(a)?.showPicker(); return 'None';
      case 'BoundingClientRect': {
        const rect = n(a)?.getBoundingClientRect();
        return result('Rect', { x: rect?.x || 0, y: rect?.y || 0,
          width: rect?.width || 0, height: rect?.height || 0 });
      }
      case 'DispatchPointer': {
        const target = doc.elementFromPoint(a.client_x, a.client_y) || doc.body;
        target.dispatchEvent(new MouseEvent(a.kind, { bubbles: true,
          clientX: a.client_x, clientY: a.client_y, button: a.button }));
        return bool(true);
      }
      case 'DispatchKeyboard': {
        (doc.activeElement || doc.body).dispatchEvent(new KeyboardEvent(a.kind, {
          bubbles: true, key: a.key, code: a.code, ctrlKey: a.ctrl_key,
          shiftKey: a.shift_key, altKey: a.alt_key, metaKey: a.meta_key
        }));
        return 'None';
      }
      default: throw new Error(`unsupported DOM operation: ${op}`);
    }
  }

  async function pump() {
    for (;;) {
      try {
        const response = await fetch('next', { cache: 'no-store' });
        if (response.status === 204) {
          await new Promise(resolve => setTimeout(resolve, 20));
          continue;
        }
        const commands = await response.json();
        const replies = [];
        for (const command of commands) {
          let resultValue = 'None', error;
          try {
            resultValue = execute(command);
            if (command.assigned_node != null) {
              const originalId = resultValue.Node;
              const node = nodes.get(originalId);
              if (!node) throw new Error('browser did not create a node');
              nodes.delete(originalId);
              nodes.set(command.assigned_node, node);
              nodeIds.set(node, command.assigned_node);
            }
          } catch (cause) {
            error = String(cause);
            if (!command.id) console.error(`Vybe ${command.operation}: ${error}`);
          }
          if (command.id) replies.push({ id: command.id, result: resultValue, error });
        }
        if (replies.length) {
          await fetch('reply', { method: 'POST', headers: { 'content-type': 'application/json' },
            body: JSON.stringify(replies) });
        }
      } catch (_) {
        await new Promise(resolve => setTimeout(resolve, 250));
      }
    }
  }
  pump();
})();
