//! A real DOM over the widget tree.
//!
//! The widgets were already most of a document: controls nest (`add_child`),
//! they answer property commands by name (`send_command_named` —
//! `SetText`/`GetText`/`SetValue`/`SetChecked`/`SetEnabled`, which is what
//! attributes and IDL properties are), and they hold every control's value.
//! What was missing is the part that makes it a DOM rather than a toolkit:
//! nodes that exist before they are inserted, an `id` attribute that is an
//! attribute rather than a rename, and events named `click`/`input`/`change`.
//! That is what this module adds, so that the engine IS the document instead
//! of having a document draped over it.
//!
//! Keeping it here rather than in the `web:*` host has a point beyond tidiness:
//! `platforms/web` declares the API and installs *a* DOM the same way it
//! installs *a* canvas backend, so the whole document can be swapped for a real
//! browser engine — or the same guest code can run in a browser, where the
//! backend is the browser's own DOM.
//!
//! **There is no second copy of any control's state.** `value`, `checked`,
//! text, enabled, visible and geometry are read back out of the widget every
//! time. What lives here is what a widget genuinely has no notion of:
//!
//! - **which nodes are in the document.** `createElement` builds a control and
//!   does NOT insert it; it renders nothing until `appendChild`. Detached
//!   controls are held here because a toolkit has nowhere to put them.
//! - **attributes with no control counterpart** — `id`, `class`, `type`,
//!   `data-*`. `id` in particular must stay an attribute: `getElementById`
//!   matches it, and it is not the widget's name.

use std::collections::HashMap;
use std::sync::{Mutex, OnceLock};

use crate::controls::make_widget;
use crate::layout::{
    find_widget_mut, take_widget, LayoutRect, PanelWidget, WidgetCommand, WidgetEvent };
use crate::{CommandValue, Form};

/// A node handle. `0` is the document itself — `document.body`, the form.
pub type NodeId = u64;

pub const DOCUMENT: NodeId = 0;

/// The bookkeeping a document owns and a control does not.
#[derive(Clone, Debug, Default)]
pub struct DomNode {
    pub id: NodeId,
    /// Lowercased tag, per `Element.tagName` normalisation for HTML.
    pub tag: String,
    /// `type` at creation, which with `tag` is what the spec says decides
    /// which control an `<input>` is. Kept because `input.type` is readable.
    pub input_type: String,
    /// Attributes that no `WidgetCommand` covers — `id`, `class`, `data-*`.
    /// Everything a control does own is forwarded to it instead.
    pub attributes: HashMap<String, String>,
    pub parent: Option<NodeId>,
    pub children: Vec<NodeId>,
}

/// One DOM event, ready for dispatch: which node, and the spec's event name.
#[derive(Clone, Debug)]
pub struct DomEvent {
    pub node: NodeId,
    /// `click`, `input`, `change`, `mouseenter`, …
    pub kind: &'static str,
}

/// A document and the widget tree that IS its rendering.
pub struct Document {
    /// The document element. A form's controls are the body's children, and
    /// `form.title` is `document.title` — not a copy of it.
    form: Form,
    nodes: HashMap<NodeId, DomNode>,
    /// Creation order, which is tree order for `getElementById`'s "first
    /// match" and for walking the document.
    order: Vec<NodeId>,
    next_id: NodeId,
    /// Created but not yet inserted. The state a toolkit has no place for.
    detached: HashMap<NodeId, Box<dyn PanelWidget>>,
}

impl Document {
    pub fn new(title: &str) -> Self {
        let mut form = Form::new(title);
        // The initial viewport. A document always has one — hit testing and
        // percentage layout both need a containing block, and `window.open`
        // overrides it from the `features` string.
        <Form as PanelWidget>::set_rect(&mut form, LayoutRect::new(0.0, 0.0, 800.0, 600.0));
        Document {
            form,
            nodes: HashMap::new(),
            order: Vec::new(),
            next_id: 0,
            detached: HashMap::new() }
    }

    /// The name the widget tree knows a node by. Internal identity, generated
    /// and unique — deliberately NOT the `id` attribute, which is the author's
    /// to set, change, duplicate or leave off.
    fn widget_name(node: NodeId) -> String {
        format!("n{}", node)
    }

    /// The body's containing block — `window.innerWidth`/`innerHeight` read
    /// back from here rather than keeping a second copy of the size.
    pub fn viewport(&self) -> LayoutRect {
        self.form.rect()
    }

    /// `window.resizeTo` / a real window resize — the viewport the body fills.
    pub fn set_viewport(&mut self, width: f32, height: f32) {
        let r = self.form.rect();
        <Form as PanelWidget>::set_rect(&mut self.form, LayoutRect::new(r.x, r.y, width, height));
    }

    pub fn form(&self) -> &Form {
        &self.form
    }

    pub fn form_mut(&mut self) -> &mut Form {
        &mut self.form
    }

    pub fn node(&self, id: NodeId) -> Option<&DomNode> {
        self.nodes.get(&id)
    }

    // ── Document ────────────────────────────────────────────────────────

    /// `document.createElement(localName)`. The control is built; it is not
    /// in the document, and renders nothing, until something appends it.
    pub fn create_element(&mut self, tag: &str, input_type: &str) -> NodeId {
        let tag = tag.trim().to_ascii_lowercase();
        let input_type = input_type.trim().to_ascii_lowercase();
        self.next_id += 1;
        let id = self.next_id;
        let kind = control_kind(&tag, &input_type);
        let (w, h) = default_size(kind);
        let mut widget = make_widget(kind, &Self::widget_name(id), "", w, h);
        // Give it its geometry immediately: `make_widget` sets a control's own
        // width/height fields, but the layout rect is what positioning and hit
        // testing use, and an unstyled element must still be visible.
        widget.set_rect(LayoutRect::new(0.0, 0.0, w, h));
        self.detached.insert(id, widget);
        self.nodes.insert(
            id,
            DomNode {
                id,
                tag,
                input_type,
                ..DomNode::default()
            },
        );
        self.order.push(id);
        id
    }

    /// `document.getElementById(elementId)` — the `id` ATTRIBUTE, first match
    /// in tree order.
    pub fn get_element_by_id(&self, element_id: &str) -> Option<NodeId> {
        self.order.iter().copied().find(|id| {
            self.nodes
                .get(id)
                .and_then(|n| n.attributes.get("id"))
                .map(|v| v == element_id)
                .unwrap_or(false)
        })
    }

    /// `document.querySelectorAll(tag)` — tag selectors only. Anything richer
    /// needs a selector engine and is not pretended here.
    pub fn elements_by_tag(&self, tag: &str) -> Vec<NodeId> {
        let tag = tag.to_ascii_lowercase();
        self.order
            .iter()
            .copied()
            .filter(|id| self.nodes.get(id).map(|n| n.tag == tag).unwrap_or(false))
            .collect()
    }

    /// `document.title` — the form's own title, read back from it.
    pub fn title(&self) -> String {
        self.form.title.clone()
    }

    pub fn set_title(&mut self, title: &str) {
        self.form.title = title.to_string();
    }

    // ── Node ────────────────────────────────────────────────────────────

    /// The unparented subtree a node currently sits in, if any. Walking up
    /// `parent` is exact — no probing — and it is what lets a node be styled
    /// or moved while its whole subtree is still outside the document.
    fn detached_root(&self, node: NodeId) -> Option<NodeId> {
        let mut cur = node;
        loop {
            match self.nodes.get(&cur)?.parent {
                None => return self.detached.contains_key(&cur).then_some(cur),
                Some(DOCUMENT) => return None,
                Some(parent) => cur = parent }
        }
    }

    /// Lift a node's control out of wherever it currently lives — the
    /// detached set, an unparented subtree, or the document — ready to be
    /// inserted somewhere else. `None` means the tree has lost it, which is a
    /// bug rather than a state a caller should paper over.
    fn extract_widget(&mut self, node: NodeId) -> Option<Box<dyn PanelWidget>> {
        if let Some(w) = self.detached.remove(&node) {
            return Some(w);
        }
        let name = Self::widget_name(node);
        match self.detached_root(node) {
            Some(root) => take_widget(self.detached.get_mut(&root)?.as_mut(), &name),
            None => take_widget(&mut self.form, &name) }
    }

    /// `parent.appendChild(child)`. Inserting is what makes an element render.
    /// A node has ONE parent, so appending it elsewhere moves it.
    ///
    /// Returns whether the child ended up in the parent. `false` is the
    /// spec's `HierarchyRequestError` case — a parent that cannot have
    /// children — surfaced rather than swallowed.
    pub fn append_child(&mut self, parent: NodeId, child: NodeId) -> bool {
        if child == DOCUMENT || child == parent || !self.nodes.contains_key(&child) {
            return false;
        }
        if parent != DOCUMENT && !self.nodes.contains_key(&parent) {
            return false;
        }
        // Appending an ancestor into its own descendant would cycle the tree.
        if parent != DOCUMENT && self.is_ancestor(child, parent) {
            return false;
        }

        let Some(widget) = self.extract_widget(child) else {
            return false;
        };

        // Unlink from the previous parent — the one-parent rule.
        if let Some(previous) = self.nodes.get(&child).and_then(|n| n.parent) {
            if let Some(p) = self.nodes.get_mut(&previous) {
                p.children.retain(|c| *c != child);
            }
        }

        let inserted = self.insert_widget(parent, widget);
        match inserted {
            None => {
                if let Some(p) = self.nodes.get_mut(&parent) {
                    p.children.push(child);
                }
                if let Some(n) = self.nodes.get_mut(&child) {
                    n.parent = Some(parent);
                }
                true
            }
            // The parent is not a container. Put the control back where it
            // can be found rather than dropping it on the floor.
            Some(widget) => {
                self.detached.insert(child, widget);
                if let Some(n) = self.nodes.get_mut(&child) {
                    n.parent = None;
                }
                false
            }
        }
    }

    /// Place a control under `parent`, wherever that parent is. Returns the
    /// control back when the parent cannot hold children.
    fn insert_widget(
        &mut self,
        parent: NodeId,
        widget: Box<dyn PanelWidget>,
    ) -> Option<Box<dyn PanelWidget>> {
        if parent == DOCUMENT {
            // The body positions absolutely — the endgame shape: an HTML page
            // whose controls carry `position:absolute` coordinates.
            let r = widget.rect();
            self.form.add_boxed_control(widget, r.x, r.y, r.w, r.h);
            return None;
        }
        let name = Self::widget_name(parent);
        let container = match self.detached_root(parent) {
            Some(root) => find_widget_mut(self.detached.get_mut(&root)?.as_mut(), &name),
            None => find_widget_mut(&mut self.form, &name) };
        match container {
            Some(c) => c.add_child(widget),
            None => Some(widget) }
    }

    fn is_ancestor(&self, ancestor: NodeId, of: NodeId) -> bool {
        let mut cur = of;
        while let Some(parent) = self.nodes.get(&cur).and_then(|n| n.parent) {
            if parent == ancestor {
                return true;
            }
            cur = parent;
        }
        false
    }

    /// `parent.removeChild(child)` — back out of the document, not destroyed.
    pub fn remove_child(&mut self, parent: NodeId, child: NodeId) -> bool {
        let Some(widget) = self.extract_widget(child) else {
            return false;
        };
        if let Some(p) = self.nodes.get_mut(&parent) {
            p.children.retain(|c| *c != child);
        }
        if let Some(n) = self.nodes.get_mut(&child) {
            n.parent = None;
        }
        self.detached.insert(child, widget);
        true
    }

    /// `element.isConnected`.
    pub fn connected(&self, node: NodeId) -> bool {
        node == DOCUMENT || self.nodes.get(&node).map(|n| n.parent.is_some()).unwrap_or(false)
    }

    // ── Commands: the one route to a control ────────────────────────────

    /// Route a command to a node's control, wherever it is — still detached,
    /// a child of the body, or nested inside a container.
    fn command(&mut self, node: NodeId, cmd: &WidgetCommand) -> CommandValue {
        let name = Self::widget_name(node);
        // A node inside an unparented subtree is still configurable — that is
        // the whole point of building an element before inserting it.
        if let Some(root) = self.detached_root(node) {
            return self
                .detached
                .get_mut(&root)
                .and_then(|w| w.send_command_named(&name, cmd))
                .unwrap_or(CommandValue::None);
        }
        self.form.send_command(&name, cmd)
    }

    fn widget_mut(&mut self, node: NodeId) -> Option<&mut dyn PanelWidget> {
        let name = Self::widget_name(node);
        match self.detached_root(node) {
            Some(root) => find_widget_mut(self.detached.get_mut(&root)?.as_mut(), &name),
            None => find_widget_mut(&mut self.form, &name) }
    }

    // ── Element: attributes ─────────────────────────────────────────────

    /// `element.setAttribute(qualifiedName, value)`.
    ///
    /// Content attributes that a control owns are forwarded to it, because
    /// that is where the state is. The rest — `id`, `class`, `data-*` — are
    /// the document's, and are kept here because nothing else has them.
    pub fn set_attribute(&mut self, node: NodeId, name: &str, value: &str) {
        let name = name.to_ascii_lowercase();
        if node == DOCUMENT {
            if name == "title" {
                self.set_title(value);
            }
            self.nodes.entry(DOCUMENT).or_default().attributes.insert(name, value.to_string());
            return;
        }
        match name.as_str() {
            "value" => self.set_value(node, value),
            "checked" => self.set_checked(node, !value.eq_ignore_ascii_case("false")),
            // Boolean content attributes: PRESENCE is truth, so
            // `disabled=""` disables. `removeAttribute` re-enables.
            "disabled" => {
                self.command(node, &WidgetCommand::SetEnabled(false));
            }
            "hidden" => {
                self.command(node, &WidgetCommand::SetVisible(false));
            }
            "min" => {
                self.command(
                    node,
                    &WidgetCommand::Custom("SetMin".into(), CommandValue::Text(value.into())),
                );
            }
            "max" => {
                self.command(
                    node,
                    &WidgetCommand::Custom("SetMax".into(), CommandValue::Text(value.into())),
                );
            }
            "placeholder" | "alt" | "title" | "src" | "href" => {
                self.command(
                    node,
                    &WidgetCommand::Custom(
                        format!("Set{}", capitalize(&name)),
                        CommandValue::Text(value.into()),
                    ),
                );
            }
            _ => {}
        }
        if let Some(n) = self.nodes.get_mut(&node) {
            n.attributes.insert(name, value.to_string());
        }
    }

    /// `element.getAttribute(qualifiedName)` — `None` when absent, which the
    /// surface turns into `null`.
    pub fn get_attribute(&mut self, node: NodeId, name: &str) -> Option<String> {
        let name = name.to_ascii_lowercase();
        // Live properties answer from the control, not from what was written.
        match name.as_str() {
            "value" => return Some(self.value(node)),
            "checked" => {
                return self.checked(node).then(|| "".to_string());
            }
            _ => {}
        }
        self.nodes.get(&node)?.attributes.get(&name).cloned()
    }

    /// `element.removeAttribute(qualifiedName)`.
    pub fn remove_attribute(&mut self, node: NodeId, name: &str) {
        let name = name.to_ascii_lowercase();
        match name.as_str() {
            "disabled" => {
                self.command(node, &WidgetCommand::SetEnabled(true));
            }
            "hidden" => {
                self.command(node, &WidgetCommand::SetVisible(true));
            }
            "checked" => self.set_checked(node, false),
            _ => {}
        }
        if let Some(n) = self.nodes.get_mut(&node) {
            n.attributes.remove(&name);
        }
    }

    // ── CSSStyleDeclaration ─────────────────────────────────────────────

    /// `element.style.setProperty(property, value)` — CSS text, units and all.
    /// Geometry lands on the widget's rect, which is where geometry lives.
    pub fn set_style_property(&mut self, node: NodeId, property: &str, value: &str) {
        let property = property.to_ascii_lowercase();
        let px = parse_px(value);
        match property.as_str() {
            "left" | "top" | "width" | "height" => {
                let Some(px) = px else { return };
                if node == DOCUMENT {
                    let r = self.form.rect();
                    let r = apply_axis(r, &property, px);
                    <Form as PanelWidget>::set_rect(&mut self.form, r);
                    return;
                }
                let Some(w) = self.widget_mut(node) else { return };
                let r = apply_axis(w.rect(), &property, px);
                w.set_rect(r);
            }
            "display" => {
                let visible = !value.eq_ignore_ascii_case("none");
                self.command(node, &WidgetCommand::SetVisible(visible));
            }
            "visibility" => {
                let visible = !value.eq_ignore_ascii_case("hidden");
                self.command(node, &WidgetCommand::SetVisible(visible));
            }
            "color" => {
                self.command(
                    node,
                    &WidgetCommand::Custom("SetForeColor".into(), CommandValue::Text(value.into())),
                );
            }
            "background-color" | "background" => {
                if node == DOCUMENT {
                    if let Some(rgba) = crate::layout::command_color(&CommandValue::Text(value.into())) {
                        self.form.background = rgba;
                    }
                    return;
                }
                self.command(
                    node,
                    &WidgetCommand::Custom("SetBackColor".into(), CommandValue::Text(value.into())),
                );
            }
            "font-size" => {
                let Some(px) = px else { return };
                self.command(
                    node,
                    &WidgetCommand::Custom(
                        "SetFontSize".into(),
                        CommandValue::Number(px as f64),
                    ),
                );
            }
            _ => {}
        }
    }

    /// The computed geometry, as CSS text — read off the control.
    pub fn style_property(&mut self, node: NodeId, property: &str) -> String {
        let property = property.to_ascii_lowercase();
        let rect = if node == DOCUMENT {
            self.form.rect()
        } else {
            match self.widget_mut(node) {
                Some(w) => w.rect(),
                None => return String::new() }
        };
        match property.as_str() {
            "left" => format!("{}px", rect.x),
            "top" => format!("{}px", rect.y),
            "width" => format!("{}px", rect.w),
            "height" => format!("{}px", rect.h),
            _ => String::new() }
    }

    // ── IDL properties ──────────────────────────────────────────────────

    /// `input.value` / `select.value` / `textarea.value`. Which command that
    /// is depends on the control, which is exactly what HTML says: `value`
    /// means different things to a text field, a range and a select.
    pub fn set_value(&mut self, node: NodeId, value: &str) {
        let cmd = match self.value_kind(node) {
            ValueKind::Text => WidgetCommand::SetText(value.to_string()),
            ValueKind::Number => WidgetCommand::SetValue(value.trim().parse().unwrap_or(0.0)),
            ValueKind::Index => {
                WidgetCommand::SetSelectedIndex(value.trim().parse().unwrap_or(0))
            }
            ValueKind::Checked => {
                WidgetCommand::SetChecked(!value.eq_ignore_ascii_case("false") && !value.is_empty())
            }
        };
        self.command(node, &cmd);
    }

    pub fn value(&mut self, node: NodeId) -> String {
        match self.command(node, &WidgetCommand::GetValue) {
            CommandValue::Text(s) => s,
            CommandValue::Number(n) => format_number(n),
            CommandValue::Index(i) => i.to_string(),
            // A checkbox's `value` is its SUBMISSION value, not its state —
            // `checked` is the state — and HTML's default for it is `"on"`
            // whether or not the box is ticked.
            CommandValue::Bool(_) => "on".to_string(),
            CommandValue::Color(r, g, b, _) => format!("#{:02x}{:02x}{:02x}", r, g, b),
            CommandValue::None => match self.command(node, &WidgetCommand::GetText) {
                CommandValue::Text(s) => s,
                _ => String::new() } }
    }

    /// `input.checked` — a boolean, read from the control.
    pub fn set_checked(&mut self, node: NodeId, checked: bool) {
        self.command(node, &WidgetCommand::SetChecked(checked));
    }

    pub fn checked(&mut self, node: NodeId) -> bool {
        matches!(self.command(node, &WidgetCommand::GetValue), CommandValue::Bool(true))
    }

    /// `node.textContent` — the control's own text.
    pub fn set_text_content(&mut self, node: NodeId, text: &str) {
        if node == DOCUMENT {
            self.set_title(text);
            return;
        }
        // A `<select>`/`<datalist>` has no text of its own: its content is its
        // options, so setting it replaces the option list.
        if self.value_kind(node) == ValueKind::Index {
            self.command(node, &WidgetCommand::ClearItems);
            for line in text.lines().filter(|l| !l.trim().is_empty()) {
                self.command(node, &WidgetCommand::AddItem(line.trim().to_string()));
            }
            return;
        }
        self.command(node, &WidgetCommand::SetText(text.to_string()));
    }

    pub fn text_content(&mut self, node: NodeId) -> String {
        if node == DOCUMENT {
            return self.title();
        }
        match self.command(node, &WidgetCommand::GetText) {
            CommandValue::Text(s) => s,
            _ => String::new() }
    }

    /// `element.focus()`.
    pub fn focus(&mut self, node: NodeId) {
        self.command(node, &WidgetCommand::Focus);
    }

    /// `select.add(option)` / `select.remove(index)`.
    pub fn add_item(&mut self, node: NodeId, text: &str) {
        self.command(node, &WidgetCommand::AddItem(text.to_string()));
    }

    pub fn remove_item(&mut self, node: NodeId, index: usize) {
        self.command(node, &WidgetCommand::RemoveItem(index));
    }

    pub fn clear_items(&mut self, node: NodeId) {
        self.command(node, &WidgetCommand::ClearItems);
    }

    /// What `value` means for this node's control.
    fn value_kind(&self, node: NodeId) -> ValueKind {
        let Some(n) = self.nodes.get(&node) else {
            return ValueKind::Text;
        };
        value_kind_of(&n.tag, &n.input_type)
    }

    // ── Events ──────────────────────────────────────────────────────────

    /// Drain what the user did into DOM events. The widget layer reports in
    /// its own vocabulary (`ButtonClicked`, `TextChanged`); this is the one
    /// place that vocabulary becomes `click` / `input` / `change`.
    pub fn drain_events(&mut self) -> Vec<DomEvent> {
        let raw = self.form.drain_events();
        let mut out = Vec::new();
        for event in raw {
            let (name, kinds): (String, &[&'static str]) = match event {
                WidgetEvent::ButtonClicked(n) => (n, &["click"]),
                WidgetEvent::LinkClicked(n) => (n, &["click"]),
                WidgetEvent::Action(n) => (n, &["click"]),
                // A checkbox/radio fires `input` then `change` together —
                // its value is committed the instant it is toggled.
                WidgetEvent::CheckboxToggled(n, _) => (n, &["input", "change"]),
                WidgetEvent::RadioSelected(n, _) => (n, &["input", "change"]),
                WidgetEvent::ColorChanged(n, _) => (n, &["input", "change"]),
                // Continuous controls fire `input` as they move.
                WidgetEvent::TextChanged(n, _) => (n, &["input"]),
                WidgetEvent::SliderChanged(n, _) => (n, &["input"]),
                WidgetEvent::NumericChanged(n, _) => (n, &["input"]),
                // A committed selection is `change`.
                WidgetEvent::SelectChanged(n, _) => (n, &["change"]),
                WidgetEvent::DropdownSelected(n, _) => (n, &["change"]),
                WidgetEvent::ListBoxSelected(n, _) => (n, &["change"]),
                WidgetEvent::ListViewSelected(n, _) => (n, &["change"]),
                WidgetEvent::TabControlChanged(n, _) => (n, &["change"]),
                WidgetEvent::CalendarDateSelected(n, _) => (n, &["change"]),
                WidgetEvent::MenuItemClicked(n, _) => (n, &["click"]),
                WidgetEvent::ToolStripItemClicked(n, _) => (n, &["click"]),
                WidgetEvent::ContextMenuItemClicked(n, _) => (n, &["click"]),
                WidgetEvent::ScrollChanged(n, _) => (n, &["scroll"]),
                WidgetEvent::MouseEnter(n) => (n, &["mouseenter"]),
                WidgetEvent::MouseLeave(n) => (n, &["mouseleave"]),
                // Container-internal events with no addressable element.
                WidgetEvent::TabChanged(_)
                | WidgetEvent::TabCloseRequested(_)
                | WidgetEvent::StatusBarClick(_)
                | WidgetEvent::TreeItemOpened(_)
                | WidgetEvent::MenuAction(_)
                | WidgetEvent::SplitMoved(_, _) => continue };
            let Some(node) = self.node_for_widget(&name) else {
                continue;
            };
            for kind in kinds {
                out.push(DomEvent { node, kind });
            }
        }
        out
    }

    /// The node a widget name belongs to. Names are `n<id>`, so this is a
    /// parse rather than a second index.
    pub fn node_for_widget(&self, widget_name: &str) -> Option<NodeId> {
        let id: NodeId = widget_name.strip_prefix('n')?.parse().ok()?;
        self.nodes.contains_key(&id).then_some(id)
    }
}

// ── The set of open documents ───────────────────────────────────────────
//
// One per browsing context. Kept here rather than in the host because the
// documents ARE the engine's state; a host holds handles to them.

/// A document handle — `window.document`.
pub type DocumentId = u64;

#[derive(Default)]
struct Documents {
    docs: HashMap<DocumentId, Mutex<Document>>,
    next_id: DocumentId,
}

fn documents() -> &'static Mutex<Documents> {
    static DOCS: OnceLock<Mutex<Documents>> = OnceLock::new();
    DOCS.get_or_init(|| Mutex::new(Documents::default()))
}

/// Open a document — the initial `about:blank` of a browsing context.
pub fn new_document(title: &str) -> DocumentId {
    let mut docs = documents().lock().unwrap();
    docs.next_id += 1;
    let id = docs.next_id;
    docs.docs.insert(id, Mutex::new(Document::new(title)));
    id
}

/// Borrow a document. `None` if the handle names no open document.
pub fn with_document<T>(id: DocumentId, f: impl FnOnce(&mut Document) -> T) -> Option<T> {
    let docs = documents().lock().unwrap();
    let doc = docs.docs.get(&id)?;
    let mut doc = doc.lock().unwrap();
    Some(f(&mut doc))
}

/// What `value` means for a given element — HTML's own split.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
enum ValueKind {
    Text,
    Number,
    Index,
    Checked,
}

fn value_kind_of(tag: &str, input_type: &str) -> ValueKind {
    match tag {
        "select" | "datalist" => ValueKind::Index,
        "progress" | "meter" => ValueKind::Number,
        "input" => match input_type {
            "checkbox" | "radio" => ValueKind::Checked,
            "range" | "number" => ValueKind::Number,
            _ => ValueKind::Text },
        _ => ValueKind::Text }
}

/// `<tag type=…>` → the control kind. This is the DOM's half of the mapping;
/// building the control from a kind is [`make_widget`]'s, shared with every
/// other surface.
fn control_kind(tag: &str, input_type: &str) -> &'static str {
    match tag {
        "input" => match input_type {
            "checkbox" => "checkbox",
            "radio" => "radiobutton",
            "range" => "trackbar",
            "number" => "numericupdown",
            "date" | "datetime-local" | "time" | "month" | "week" => "datetimepicker",
            "button" | "submit" | "reset" => "button",
            "image" => "picturebox",
            // `text`, `password`, `email`, `search`, `tel`, `url`, and the
            // spec's "missing value default" for anything unknown.
            _ => "textbox" },
        "button" => "button",
        "select" => "combobox",
        "datalist" => "listbox",
        "textarea" => "richtextbox",
        "progress" | "meter" => "progressbar",
        "a" => "linklabel",
        "img" => "picturebox",
        "canvas" => "canvas",
        // A block container. `Panel`/`GroupBox` are backgrounds that hold no
        // children, so a nestable container is the only honest mapping — and
        // vertical flow IS what CSS `display:block` does to children.
        "fieldset" => "flowlayoutpanel",
        "table" => "datagridview",
        "ul" | "ol" => "listbox",
        "menu" => "menustrip",
        "dialog" | "div" | "form" | "section" | "article" | "main" | "aside" | "header"
        | "footer" | "nav" | "li" => "flowlayoutpanel",
        // Anything else is text-bearing: `p`, `span`, `label`, `h1`…`h6`, `li`.
        _ => "label" }
}

/// A starting size, so a control that is appended before it is styled is
/// visible rather than zero-sized. CSS overrides it.
fn default_size(kind: &str) -> (f32, f32) {
    match kind {
        "textbox" | "richtextbox" | "combobox" | "numericupdown" | "datetimepicker" => (160.0, 24.0),
        "button" => (96.0, 28.0),
        "checkbox" | "radiobutton" => (140.0, 20.0),
        "trackbar" | "progressbar" => (160.0, 20.0),
        "listbox" | "listview" | "datagridview" | "treeview" => (200.0, 120.0),
        "panel" | "groupbox" | "flowlayoutpanel" | "canvas" | "picturebox" => (200.0, 150.0),
        _ => (120.0, 20.0) }
}

/// `"12px"` / `"12"` / `"1.5em"` → pixels. Unitless and `px` are exact; `em`
/// and `rem` use the 16px root default. Percentages need a containing block
/// and are left to the layout containers.
fn parse_px(value: &str) -> Option<f32> {
    let v = value.trim();
    if let Some(n) = v.strip_suffix("px") {
        return n.trim().parse().ok();
    }
    if let Some(n) = v.strip_suffix("em").or_else(|| v.strip_suffix("rem")) {
        return n.trim().parse::<f32>().ok().map(|n| n * 16.0);
    }
    if let Some(n) = v.strip_suffix("pt") {
        return n.trim().parse::<f32>().ok().map(|n| n * 4.0 / 3.0);
    }
    v.parse().ok()
}

fn apply_axis(r: LayoutRect, property: &str, px: f32) -> LayoutRect {
    match property {
        "left" => LayoutRect::new(px, r.y, r.w, r.h),
        "top" => LayoutRect::new(r.x, px, r.w, r.h),
        "width" => LayoutRect::new(r.x, r.y, px, r.h),
        "height" => LayoutRect::new(r.x, r.y, r.w, px),
        _ => r }
}

/// JS number formatting for an IDL `value`: integers carry no `.0`.
fn format_number(n: f64) -> String {
    if n.fract() == 0.0 && n.abs() < 1e21 {
        format!("{}", n as i64)
    } else {
        format!("{}", n)
    }
}

fn capitalize(s: &str) -> String {
    let mut c = s.chars();
    match c.next() {
        None => String::new(),
        Some(f) => f.to_uppercase().collect::<String>() + c.as_str() }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn created_element_is_not_in_the_document() {
        let mut doc = Document::new("t");
        let cb = doc.create_element("input", "checkbox");
        assert!(!doc.connected(cb), "createElement must not insert");
        assert_eq!(doc.form().control_count(), 0);
        doc.append_child(DOCUMENT, cb);
        assert!(doc.connected(cb));
        assert_eq!(doc.form().control_count(), 1);
    }

    #[test]
    fn checked_is_the_widgets_state_not_a_copy() {
        let mut doc = Document::new("t");
        let cb = doc.create_element("input", "checkbox");
        doc.append_child(DOCUMENT, cb);
        doc.set_checked(cb, true);
        assert!(doc.checked(cb));
        // Read it straight off the control to prove there is no second copy.
        let v = doc.form_mut().send_command("n1", &WidgetCommand::GetValue);
        assert!(matches!(v, CommandValue::Bool(true)));
    }

    #[test]
    fn value_survives_being_set_before_insertion() {
        let mut doc = Document::new("t");
        let input = doc.create_element("input", "text");
        doc.set_value(input, "hello");
        assert_eq!(doc.value(input), "hello");
        doc.append_child(DOCUMENT, input);
        assert_eq!(doc.value(input), "hello");
    }

    #[test]
    fn id_is_an_attribute_not_a_rename() {
        let mut doc = Document::new("t");
        let b = doc.create_element("button", "");
        doc.set_attribute(b, "id", "ok");
        assert_eq!(doc.get_element_by_id("ok"), Some(b));
        assert_eq!(doc.get_attribute(b, "id").as_deref(), Some("ok"));
        assert_eq!(doc.get_attribute(b, "class"), None, "absent attribute is null");
    }

    #[test]
    fn a_node_has_one_parent() {
        let mut doc = Document::new("t");
        let a = doc.create_element("div", "");
        let b = doc.create_element("div", "");
        let child = doc.create_element("button", "");
        doc.append_child(DOCUMENT, a);
        doc.append_child(DOCUMENT, b);
        doc.append_child(a, child);
        doc.append_child(b, child);
        assert_eq!(doc.node(a).unwrap().children, Vec::<NodeId>::new());
        assert_eq!(doc.node(b).unwrap().children, vec![child]);
    }

    #[test]
    fn a_subtree_can_be_built_before_it_is_inserted() {
        // The canonical DOM build order: assemble off-document, insert once.
        let mut doc = Document::new("t");
        let panel = doc.create_element("div", "");
        let field = doc.create_element("input", "text");
        assert!(doc.append_child(panel, field), "nesting off-document");
        // Reachable and configurable while the whole subtree is detached.
        doc.set_value(field, "typed");
        assert_eq!(doc.value(field), "typed");
        doc.set_style_property(field, "width", "70px");
        assert_eq!(doc.style_property(field, "width"), "70px");

        assert!(doc.append_child(DOCUMENT, panel));
        assert!(doc.connected(field), "a child of an inserted node is connected");
        assert_eq!(doc.value(field), "typed", "state survives insertion");
    }

    #[test]
    fn a_child_moves_out_of_a_nested_container() {
        let mut doc = Document::new("t");
        let outer = doc.create_element("div", "");
        let inner = doc.create_element("div", "");
        let b = doc.create_element("button", "");
        doc.append_child(DOCUMENT, outer);
        assert!(doc.append_child(outer, inner));
        assert!(doc.append_child(inner, b));
        doc.set_value(b, "deep");

        // Moving it to the body must find it two levels down and take it.
        assert!(doc.append_child(DOCUMENT, b), "move out of a nested container");
        assert_eq!(doc.node(inner).unwrap().children, Vec::<NodeId>::new());
        assert_eq!(doc.value(b), "deep", "the same control, not a rebuild");
    }

    #[test]
    fn appending_to_a_leaf_reports_failure() {
        // A `<button>` is not a container. The spec throws
        // HierarchyRequestError; this reports it rather than dropping the
        // control silently.
        let mut doc = Document::new("t");
        let b = doc.create_element("button", "");
        let label = doc.create_element("span", "");
        doc.append_child(DOCUMENT, b);
        assert!(!doc.append_child(b, label), "a leaf cannot take children");
        assert!(!doc.connected(label));
        // Still usable — not lost.
        doc.set_text_content(label, "still here");
        assert_eq!(doc.text_content(label), "still here");
    }

    #[test]
    fn a_node_cannot_contain_itself() {
        let mut doc = Document::new("t");
        let outer = doc.create_element("div", "");
        let inner = doc.create_element("div", "");
        doc.append_child(DOCUMENT, outer);
        doc.append_child(outer, inner);
        assert!(!doc.append_child(inner, outer), "cycles must be refused");
        assert!(doc.connected(inner), "the refused move changes nothing");
    }

    #[test]
    fn removing_a_nested_child_takes_it_out_of_the_document() {
        let mut doc = Document::new("t");
        let panel = doc.create_element("div", "");
        let b = doc.create_element("button", "");
        doc.append_child(DOCUMENT, panel);
        doc.append_child(panel, b);
        assert!(doc.remove_child(panel, b));
        assert!(!doc.connected(b));
        // Re-insertable, because removeChild detaches rather than destroys.
        assert!(doc.append_child(DOCUMENT, b));
        assert!(doc.connected(b));
    }

    #[test]
    fn style_geometry_lands_on_the_control() {
        let mut doc = Document::new("t");
        let b = doc.create_element("button", "");
        doc.append_child(DOCUMENT, b);
        doc.set_style_property(b, "left", "10px");
        doc.set_style_property(b, "top", "20px");
        doc.set_style_property(b, "width", "80px");
        assert_eq!(doc.style_property(b, "left"), "10px");
        assert_eq!(doc.style_property(b, "top"), "20px");
        assert_eq!(doc.style_property(b, "width"), "80px");
    }

    #[test]
    fn a_click_arrives_as_a_dom_event() {
        use crate::layout::{MouseButton, MouseEvent, MouseEventKind};
        let mut doc = Document::new("t");
        let b = doc.create_element("button", "");
        doc.append_child(DOCUMENT, b);
        doc.set_style_property(b, "left", "0px");
        doc.set_style_property(b, "top", "0px");
        let form = doc.form_mut();
        let press = MouseEvent {
            kind: MouseEventKind::Press(MouseButton::Left),
            x: 10.0,
            y: 10.0,
            cmd: false,
            shift: false,
            alt: false };
        form.handle_mouse(&press);
        form.handle_mouse(&MouseEvent {
            kind: MouseEventKind::Release(MouseButton::Left),
            ..press
        });
        let events = doc.drain_events();
        assert!(
            events.iter().any(|e| e.node == b && e.kind == "click"),
            "expected a click on the button, got {:?}",
            events
        );
    }
}
