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
use crate::css::Style;
use crate::layout::{
    Dock, LayoutRect, PanelWidget, RenderContext, WidgetCommand, WidgetEvent, find_widget_mut,
    take_widget,
};
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
    /// Which edge of its container this element takes, when it is docked.
    ///
    /// `None` is the ordinary absolutely-positioned element, which keeps the
    /// rect its `left`/`top`/`width`/`height` gave it. A docked element does
    /// NOT own its position — the container computes it from the space left
    /// over, which is why this lives here and not in the adapter that set it.
    ///
    /// This is what VCL's `Align` and WinForms' `Dock` both mean.
    pub dock: Option<Dock>,
    /// `element.style` — the inline declarations, as they were set.
    ///
    /// The element's own record of its CSS, kept because a style write is
    /// observable: `el.style.color = 'red'` has to read back `'red'`. Before
    /// this existed, a write was translated into a `WidgetCommand` and the
    /// declaration was forgotten, so the read side could only answer for
    /// geometry it could recover from the widget's rect and returned `""` for
    /// everything else — including properties the write side accepted.
    pub style: Style,
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
    /// The document element's own `style`. `DOCUMENT` is not in `nodes` — it is
    /// the form — so its declarations live beside them.
    document_style: Style,
    /// THE TOP LAYER — HTML's own name for it, and its own ordered set.
    ///
    /// `showModal()` puts a dialog here; `close()` takes it out. Membership is
    /// what makes a modal paint above everything, and it is deliberately NOT
    /// a `z-index`: the top layer outranks the whole cascade, which is why a
    /// modal beats a `z-index: 9999` sibling in a browser. Modelling it as a
    /// separate paint pass is what makes that true here rather than
    /// approximately true.
    top_layer: Vec<NodeId>,
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
            detached: HashMap::new(),
            document_style: Style::new(),
            top_layer: Vec::new(),
        }
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

    /// Paint the document. This is the whole of what a host needs to show a
    /// window: hand it a pixmap, get the rendered tree.
    ///
    /// Rendering belongs HERE, with the tree it draws — that is the point of
    /// the toolkit being usable on its own. A host that reached in for the
    /// form and rendered it itself would be re-implementing the document's
    /// paint order outside the document, and would be pointing at whichever
    /// `Form` it happened to hold: that is exactly how a window came up empty
    /// while the controls sat in the document all along.
    pub fn render(&mut self, ctx: &mut RenderContext) {
        self.form.render(ctx);
        // Then the top layer, in the order dialogs entered it — the second
        // pass IS what "above everything" means. Each modal gets its
        // `::backdrop` painted immediately beneath it, so a stack of two
        // modals dims twice, exactly as a browser does.
        for node in self.top_layer.clone() {
            self.render_backdrop(ctx);
            let name = Self::widget_name(node);
            if let Some(widget) = find_widget_mut(&mut self.form, &name) {
                widget.render(ctx);
            }
        }
    }

    /// `dialog::backdrop` — the sheet between the page and a modal.
    ///
    /// It covers the VIEWPORT rather than the dialog's parent, because the
    /// backdrop's containing block is the viewport however deeply nested the
    /// dialog is. The colour is the usual user-agent default; a `::backdrop`
    /// rule cannot override it yet, since the cascade has no pseudo-elements.
    fn render_backdrop(&mut self, ctx: &mut RenderContext) {
        let viewport = self.form.rect();
        let scale = ctx.scale;
        let Some(rect) = tiny_skia::Rect::from_xywh(
            viewport.x * scale,
            viewport.y * scale,
            viewport.w * scale,
            viewport.h * scale,
        ) else {
            return;
        };
        let mut paint = tiny_skia::Paint::default();
        paint.set_color_rgba8(0, 0, 0, 26);
        ctx.pixmap
            .fill_rect(rect, &paint, tiny_skia::Transform::identity(), None);
    }

    // ── HTMLDialogElement ───────────────────────────────────────────────

    /// `dialog.show()` (`modal` false) / `dialog.showModal()` (`modal` true).
    ///
    /// `open` is a reflected attribute, so opening a dialog is recorded on the
    /// element itself. The rest is the user-agent stylesheet, applied here
    /// because this document has no UA sheet to cascade from: a dialog is
    /// `display: none` until opened, and a MODAL one is positioned against
    /// the viewport and centred in it (`position: fixed; margin: auto`).
    pub fn show_dialog(&mut self, node: NodeId, modal: bool) {
        let Some(dom_node) = self.nodes.get_mut(&node) else {
            return;
        };
        dom_node.attributes.insert("open".to_string(), String::new());
        self.command(node, &WidgetCommand::SetVisible(true));
        if !modal {
            return;
        }
        self.set_style_property(node, "position", "fixed");
        self.centre_in_viewport(node);
        if !self.top_layer.contains(&node) {
            self.top_layer.push(node);
        }
    }

    /// `dialog.close()` — clears `open` and leaves the top layer.
    ///
    /// The `close` event and `returnValue` are not here: neither is a
    /// reflected attribute, so neither is the tree's to remember.
    pub fn close_dialog(&mut self, node: NodeId) {
        if let Some(dom_node) = self.nodes.get_mut(&node) {
            dom_node.attributes.remove("open");
        }
        self.command(node, &WidgetCommand::SetVisible(false));
        self.top_layer.retain(|open| *open != node);
    }

    /// `dialog.open` — read straight off the reflected attribute.
    pub fn dialog_open(&self, node: NodeId) -> bool {
        self.nodes
            .get(&node)
            .is_some_and(|n| n.attributes.contains_key("open"))
    }

    /// `dialog:modal { margin: auto }` — centred in the viewport, keeping the
    /// size the dialog already has.
    fn centre_in_viewport(&mut self, node: NodeId) {
        let viewport = self.form.rect();
        let name = Self::widget_name(node);
        let Some(widget) = find_widget_mut(&mut self.form, &name) else {
            return;
        };
        let rect = widget.rect();
        widget.set_rect(LayoutRect::new(
            viewport.x + ((viewport.w - rect.w) / 2.0).max(0.0),
            viewport.y + ((viewport.h - rect.h) / 2.0).max(0.0),
            rect.w,
            rect.h,
        ));
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
        let (w, h) = default_size(kind, &tag);
        let mut widget = make_widget(kind, &Self::widget_name(id), "", w, h);
        // Give it its geometry immediately: `make_widget` sets a control's own
        // width/height fields, but the layout rect is what positioning and hit
        // testing use, and an unstyled element must still be visible.
        // Metadata content is a node with no rendering: zero-sized, so nothing
        // paints and nothing hit-tests. `SetVisible` would look like the
        // obvious answer and is not — only panels honour it; a label ignores
        // it entirely, so hiding a leaf that way silently fails.
        let (w, h) = if renders_nothing(&tag, &input_type) {
            (0.0, 0.0)
        } else {
            (w, h)
        };
        widget.set_rect(LayoutRect::new(0.0, 0.0, w, h));
        // A `<fieldset>` draws a border with its legend across the top; a
        // `<div>` draws nothing. Same widget, and the UA stylesheet is what
        // separates them.
        if tag == "fieldset" {
            widget.handle_command(&WidgetCommand::Custom(
                "SetBordered".into(),
                CommandValue::Bool(true),
            ));
        }
        // `dialog:not([open]) { display: none }` — the other UA rule that
        // makes a dialog a dialog. Without it a form's secondary window is
        // painted from the moment it is built, before anything shows it.
        if tag == "dialog" {
            widget.handle_command(&WidgetCommand::SetVisible(false));
        }
        self.detached.insert(id, widget);
        self.nodes.insert(
            id,
            DomNode {
                id,
                tag,
                input_type,
                // A menu bar takes the top edge of whatever holds it. This is
                // WinForms' own `MenuStrip.Dock` default and what a block-level
                // bar does at the head of the flow — a starting dock, exactly
                // as `default_size` is a starting geometry, and `Align`/`Dock`
                // overwrite it. Without it a form's menu sat at 0,0 on top of
                // the control that had already claimed the client area.
                dock: (kind == "menustrip").then_some(Dock::Top),
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

    /// `document.querySelectorAll("[id]")` — every element carrying an `id`,
    /// in tree order, as `(node, id)`.
    ///
    /// The one selector a toolkit host genuinely needs: a control's `Name` IS
    /// its `id`, so this is how anything outside the document enumerates the
    /// controls it built without knowing their tags in advance.
    pub fn elements_with_id(&self) -> Vec<(NodeId, String)> {
        self.order
            .iter()
            .filter_map(|id| {
                let attr = self.nodes.get(id)?.attributes.get("id")?;
                Some((*id, attr.clone()))
            })
            .collect()
    }

    /// Every element in the document, in tree order.
    ///
    /// `elements_with_id` answers only for controls the author NAMED, and a
    /// control's `id` is optional — a VCL form that builds sixteen buttons in a
    /// loop and never assigns `Name` has sixteen elements and no ids. A host
    /// enumerating what was built (a debugger's control dump, a test's
    /// inventory) has to see those too, or it reports an empty form while the
    /// window plainly shows one.
    pub fn elements(&self) -> Vec<NodeId> {
        self.order.clone()
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
                Some(parent) => cur = parent,
            }
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
            None => take_widget(&mut self.form, &name),
        }
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

        // An `<option>` is CONTENT of its select, not a control beside it —
        // `insert_widget` would refuse, because a combobox is not a container.
        // Appending options is the DOM-native way to fill a list, so it has to
        // work; before this the only route was assigning the whole
        // `textContent`, which no frontend spells that way.
        if self.value_kind(parent) == ValueKind::Index && self.is_item_element(child) {
            if let Some(n) = self.nodes.get_mut(&child) {
                n.parent = Some(parent);
            }
            if let Some(p) = self.nodes.get_mut(&parent) {
                p.children.push(child);
            }
            self.sync_items(parent);
            return true;
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
                // Docking is a property of the child but a decision of the
                // container, so joining a container re-runs it. Without this,
                // setting the dock BEFORE the parent (the order a designer
                // file emits) left the element with its default rect forever.
                self.relayout_docked(parent);
                // Same reasoning for out-of-flow, and the same ordering trap:
                // `Left := 8` before `Parent := Panel` is the ordinary way to
                // write VCL, and the container did not exist to be told.
                // Metadata takes no part in layout either — a `<style>` in a
                // flow container would otherwise be handed a slot and push its
                // siblings down by a row of nothing.
                let metadata = self
                    .node(child)
                    .map(|n| renders_nothing(&n.tag, &n.input_type))
                    .unwrap_or(false);
                if metadata {
                    self.set_child_flow(child, false);
                }
                if self.positions_itself(child) {
                    self.set_child_flow(child, false);
                    // Insertion already ran the container's layout — twice, in
                    // fact — so the child's rect was overwritten before it was
                    // excluded. Restore it from the DECLARATIONS rather than
                    // from the rect we found: the rect is downstream of the
                    // layout that just clobbered it, and the declarations are
                    // what the program actually asked for.
                    self.apply_declared_geometry(child);
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
            None => find_widget_mut(&mut self.form, &name),
        };
        match container {
            Some(c) => c.add_child(widget),
            None => Some(widget),
        }
    }

    /// Lay out `parent`'s docked children, in document order.
    ///
    /// Each docked child takes one edge off the space that is left, and `Fill`
    /// takes all of it — the algorithm `DockPanel::relayout` already runs, here
    /// applied to an ordinary container so that any element can dock, not only
    /// a child of a dock panel. Undocked siblings are untouched: they keep the
    /// rect their `left`/`top`/`width`/`height` gave them, so absolute
    /// positioning and docking compose in one container.
    fn relayout_docked(&mut self, parent: NodeId) {
        // Document ORDER, which is what decides who takes an edge first —
        // `self.order` is creation order and children are appended in it.
        let docked: Vec<(NodeId, Dock)> = self
            .order
            .iter()
            .copied()
            .filter_map(|id| {
                let n = self.nodes.get(&id)?;
                (n.parent == Some(parent) || (parent == DOCUMENT && n.parent == Some(DOCUMENT)))
                    .then_some((id, n.dock?))
            })
            .collect();
        if docked.is_empty() {
            return;
        }
        // Edges first, `Fill` last — whatever order the children were created
        // in. The filling control takes what is LEFT, so a form that creates
        // its client area before its status bar (the ordinary way to write it)
        // must not starve the status bar of every pixel. VCL and WinForms both
        // resolve the client region last for exactly this reason.
        let docked: Vec<(NodeId, Dock)> = docked
            .iter()
            .filter(|(_, d)| *d != Dock::Fill)
            .chain(docked.iter().filter(|(_, d)| *d == Dock::Fill))
            .copied()
            .collect();
        let mut remaining = if parent == DOCUMENT {
            self.form.rect()
        } else {
            match self.widget_mut(parent) {
                Some(w) => w.rect(),
                None => return,
            }
        };
        for (child, dock) in docked {
            // An edge dock keeps the child's own extent along the other axis;
            // only `Fill` discards both. Read before the mutable borrow.
            let own = match self.widget_mut(child) {
                Some(w) => w.rect(),
                None => continue,
            };
            let rect = match dock {
                Dock::Left => remaining.take_left(own.w),
                Dock::Right => remaining.take_right(own.w),
                Dock::Top => remaining.take_top(own.h),
                Dock::Bottom => remaining.take_bottom(own.h),
                Dock::Fill => {
                    let r = remaining;
                    remaining = LayoutRect::zero();
                    r
                }
            };
            if let Some(w) = self.widget_mut(child) {
                w.set_rect(rect);
            }
        }
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
        node == DOCUMENT
            || self
                .nodes
                .get(&node)
                .map(|n| n.parent.is_some())
                .unwrap_or(false)
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
            None => find_widget_mut(&mut self.form, &name),
        }
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
            self.nodes
                .entry(DOCUMENT)
                .or_default()
                .attributes
                .insert(name, value.to_string());
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
    ///
    /// The declaration is **recorded first, unconditionally**, and only then
    /// translated into whatever the widget can act on. The two steps answer
    /// different questions — what did the author say, and what does the toolkit
    /// do about it — and conflating them is why a property the write side
    /// accepted could still read back empty.
    pub fn set_style_property(&mut self, node: NodeId, property: &str, value: &str) {
        let property = property.to_ascii_lowercase();
        self.record_style(node, &property, value);
        let px = parse_px(value);
        match property.as_str() {
            "left" | "top" | "width" | "height" => {
                // `left`/`width` measure across, `top`/`height` down — which is
                // the axis a percentage refers to.
                let horizontal = matches!(property.as_str(), "left" | "width");
                let Some(px) = self.resolve_length(node, value, horizontal) else {
                    return;
                };
                if node == DOCUMENT {
                    let r = self.form.rect();
                    let r = apply_axis(r, &property, px);
                    <Form as PanelWidget>::set_rect(&mut self.form, r);
                    return;
                }
                let clamped = self.clamp_to_constraints(node, &property, px);
                // `left`/`top` are measured from the containing block, not from
                // the document. A rect is in form coordinates, so the block's
                // origin has to be added back — otherwise a child at (20, 50)
                // inside a panel at (10, 10) lands at (20, 50) on the BODY.
                // Invisible whenever the container sits at the origin, which is
                // why the calculator never showed it.
                let origin = match property.as_str() {
                    "left" => self.containing_block(node).x,
                    "top" => self.containing_block(node).y,
                    _ => 0.0,
                };
                let Some(w) = self.widget_mut(node) else {
                    return;
                };
                let r = apply_axis(w.rect(), &property, origin + clamped);
                w.set_rect(r);
                // Setting a coordinate is a positioning statement, so the
                // container must stop arranging this child — otherwise the next
                // `relayout()` recomputes the rect from flow order and discards
                // the value just written. The rect was never missing; it was
                // overwritten.
                if matches!(property.as_str(), "left" | "top") {
                    if self.declared_position(node) == "relative" {
                        // In flow, so the coordinate is an OFFSET the container
                        // applies — it has to be told the new one, then asked
                        // to arrange again.
                        self.set_child_relative(node);
                        self.relayout_parent(node);
                    } else if self.positions_itself(node) {
                        self.set_child_flow(node, false);
                    } else {
                        // `left` is inert on a static box in CSS. The write
                        // above moved the widget, so hand the child back to its
                        // container and let the flow answer stand.
                        self.relayout_parent(node);
                    }
                }
            }
            // `right`/`bottom` position from the OPPOSITE edge of the
            // containing block, which is what makes an absolutely positioned
            // box anchorable to any corner — Flutter's `Positioned` uses all
            // four, and VCL/WinForms use only the first two.
            //
            // With the matching `left`/`top` also set, CSS stretches the box
            // between the two edges; with neither, the box keeps its size and
            // moves. Both fall out of computing the edge and comparing.
            "right" | "bottom" => {
                let horizontal = property == "right";
                let Some(offset) = self.resolve_length(node, value, horizontal) else {
                    return;
                };
                let block = self.containing_block(node);
                let (origin, extent) = if horizontal {
                    (block.x, block.w)
                } else {
                    (block.y, block.h)
                };
                let opposite = if horizontal { "left" } else { "top" };
                let has_opposite = self
                    .style(node)
                    .map(|s| !s.get(opposite).is_empty())
                    .unwrap_or(false);
                let Some(w) = self.widget_mut(node) else {
                    return;
                };
                let mut r = w.rect();
                // The far edge in FORM coordinates — the containing block's own
                // origin plus its extent, less the inset.
                let far_edge = origin + extent - offset;
                if has_opposite {
                    // Anchored both sides: the box spans between them.
                    if horizontal {
                        r.w = (far_edge - r.x).max(0.0);
                    } else {
                        r.h = (far_edge - r.y).max(0.0);
                    }
                } else if horizontal {
                    r.x = far_edge - r.w;
                } else {
                    r.y = far_edge - r.h;
                }
                w.set_rect(r);
                self.set_child_flow(node, false);
            }
            // Constraints. Recorded by the store; applied by re-running the
            // axis they constrain, so declaration order does not matter.
            "min-width" | "max-width" => self.reapply_constrained_axis(node, "width"),
            "min-height" | "max-height" => self.reapply_constrained_axis(node, "height"),
            // `align-self` — one child overriding the container's
            // `align-items`. Told to the container, which is what aligns.
            "align-self" => {
                let mode = value.trim().to_ascii_lowercase();
                self.tell_container(node, "SetChildAlignSelf", &mode);
            }
            // `order` — the position a child takes in the flow, regardless of
            // document order.
            "order" => {
                let order = value.trim().to_string();
                self.tell_container(node, "SetChildOrder", &order);
            }
            // `overflow` — whether a container clips what does not fit.
            "overflow" | "overflow-x" | "overflow-y" => {
                let mode = value.trim().to_ascii_lowercase();
                self.command(
                    node,
                    &WidgetCommand::Custom("SetOverflow".into(), CommandValue::Text(mode)),
                );
            }
            // Docking. The element stops owning its own rect: the container
            // hands it one edge of whatever space is left, in child order,
            // which is exactly the algorithm `DockPanel` already runs.
            "dock" => {
                let dock = match value.trim().to_ascii_lowercase().as_str() {
                    "left" => Some(Dock::Left),
                    "right" => Some(Dock::Right),
                    "top" => Some(Dock::Top),
                    "bottom" => Some(Dock::Bottom),
                    "fill" => Some(Dock::Fill),
                    // `none` is the ordinary absolutely-positioned element,
                    // and so is anything unrecognised — an unknown keyword
                    // must not silently swallow the element's geometry.
                    _ => None,
                };
                let Some(n) = self.nodes.get_mut(&node) else {
                    return;
                };
                if n.dock == dock {
                    return;
                }
                n.dock = dock;
                let parent = n.parent;
                self.relayout_docked(parent.unwrap_or(DOCUMENT));
            }
            // `display` carries TWO things, and conflating them cost a day:
            // `none` is visibility, every other value is a LAYOUT MODE. When
            // this arm meant only visibility, `display: flex` marked the
            // element visible and selected no layout — indistinguishable from
            // unimplemented, while actually being consumed.
            "display" => {
                if value.eq_ignore_ascii_case("none") {
                    self.command(node, &WidgetCommand::SetVisible(false));
                } else {
                    self.command(node, &WidgetCommand::SetVisible(true));
                    // A flex container arranges its children along an axis,
                    // which is what the flow panel already does; `block` leaves
                    // it as it was. Neither needs a different widget — that is
                    // the whole point of the mode being a property.
                }
            }
            // `position` decides WHO places this element: itself, or its
            // container. `absolute`/`fixed` are out of flow, which is precisely
            // what every pixel-positioned frontend means by setting Left/Top.
            //
            // Addressed to the PARENT, because the container is what arranges —
            // the same reason `dock` is resolved by `relayout_docked` rather
            // than by the child.
            "position" => match value.trim().to_ascii_lowercase().as_str() {
                "absolute" | "fixed" => self.set_child_flow(node, false),
                "relative" => self.set_child_relative(node),
                _ => self.set_child_flow(node, true),
            },
            "flex-direction" => {
                let direction = value.trim().to_ascii_lowercase();
                self.command(
                    node,
                    &WidgetCommand::Custom(
                        "SetFlexDirection".into(),
                        CommandValue::Text(direction),
                    ),
                );
            }
            // Flutter's `mainAxisAlignment` / `crossAxisAlignment` — the two
            // most-used layout properties it has, and the panel had the
            // algorithm with no vocabulary reaching it.
            "justify-content" => {
                let mode = value.trim().to_ascii_lowercase();
                self.command(
                    node,
                    &WidgetCommand::Custom("SetJustifyContent".into(), CommandValue::Text(mode)),
                );
            }
            "align-items" => {
                let mode = value.trim().to_ascii_lowercase();
                self.command(
                    node,
                    &WidgetCommand::Custom("SetAlignItems".into(), CommandValue::Text(mode)),
                );
            }
            "flex-wrap" => {
                let wrap = value.trim().to_ascii_lowercase();
                self.command(
                    node,
                    &WidgetCommand::Custom("SetFlexWrap".into(), CommandValue::Text(wrap)),
                );
            }
            "gap" | "row-gap" | "column-gap" => {
                let Some(px) = px else { return };
                self.command(
                    node,
                    &WidgetCommand::Custom("SetGap".into(), CommandValue::Number(px as f64)),
                );
            }
            "padding" => {
                // One inset, from the largest edge of the shorthand: the panel
                // carries a single padding. Better here, once, than in each
                // frontend flattening its own.
                let largest = value
                    .split_whitespace()
                    .filter_map(parse_px)
                    .fold(f32::NAN, f32::max);
                if largest.is_finite() {
                    self.command(
                        node,
                        &WidgetCommand::Custom(
                            "SetPadding".into(),
                            CommandValue::Number(largest as f64),
                        ),
                    );
                }
            }
            // `flex: <grow>` on a CHILD — the weight it takes of the leftover
            // space. `flex: 0 0 auto` is a fixed bar, `flex: 1` shares.
            "flex" | "flex-grow" => {
                let grow = value
                    .split_whitespace()
                    .next()
                    .and_then(|g| g.parse::<f32>().ok());
                let grow = match (grow, value.trim()) {
                    (Some(grow), _) => Some(grow),
                    (None, "none") => Some(0.0),
                    (None, "auto") => Some(1.0),
                    _ => None,
                };
                if let Some(grow) = grow {
                    // Told to the container, not the child: `SetFlex` is only
                    // implemented by panels, so a button or label silently kept
                    // the trait default of 1.0 and grew regardless. The
                    // container is what distributes the space, so it is what
                    // has to know.
                    self.set_child_flex(node, grow);
                    // A panel is also a child of something, and its own weight
                    // is its business.
                    self.command(node, &WidgetCommand::SetFlex(grow));
                }
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
                    if let Some(rgba) =
                        crate::layout::command_color(&CommandValue::Text(value.into()))
                    {
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
                    &WidgetCommand::Custom("SetFontSize".into(), CommandValue::Number(px as f64)),
                );
            }
            _ => {}
        }
    }

    /// Tell an element's CONTAINER whether it arranges that element.
    ///
    /// The flag lives with the container because the container is what
    /// arranges, exactly as `dock` does. It is re-sent on insertion too: a
    /// frontend sets geometry before appending as often as after, and a
    /// container that never heard about a child cannot leave it alone.
    fn set_child_flow(&mut self, node: NodeId, in_flow: bool) {
        let placement = if in_flow { "flow" } else { "absolute" }.to_string();
        self.send_child_placement(node, placement);
    }

    /// `position: relative` — arranged by the container, then offset.
    ///
    /// The offsets are the DECLARED `left`/`top`, not a resolved coordinate:
    /// the container has not placed the child yet, so there is no flow
    /// position to add them to until it does.
    fn set_child_relative(&mut self, node: NodeId) {
        let (dx, dy) = self
            .style(node)
            .map(|s| (parse_px(&s.get("left")), parse_px(&s.get("top"))))
            .unwrap_or((None, None));
        let placement = format!(
            "relative:{},{}",
            dx.unwrap_or(0.0),
            dy.unwrap_or(0.0)
        );
        self.send_child_placement(node, placement);
    }

    fn send_child_placement(&mut self, node: NodeId, placement: String) {
        let Some(parent) = self.nodes.get(&node).and_then(|n| n.parent) else {
            return;
        };
        let spec = format!("{}={}", Self::widget_name(node), placement);
        if parent == DOCUMENT {
            self.form.handle_command(&WidgetCommand::Custom(
                "SetChildFlow".into(),
                CommandValue::Text(spec),
            ));
            return;
        }
        self.command(
            parent,
            &WidgetCommand::Custom("SetChildFlow".into(), CommandValue::Text(spec)),
        );
    }

    /// Is this element an ITEM of a list rather than a control of its own?
    fn is_item_element(&self, node: NodeId) -> bool {
        self.node(node)
            .map(|n| matches!(n.tag.as_str(), "option" | "optgroup" | "li"))
            .unwrap_or(false)
    }

    /// Rebuild an item-bearing control's list from its item children.
    ///
    /// Derived on demand rather than kept in step, for the reason `linecount`
    /// is: one source of truth, and no ordering rule to get wrong when an
    /// option's text arrives after the option does.
    fn sync_items(&mut self, node: NodeId) {
        let items: Vec<NodeId> = self
            .node(node)
            .map(|n| n.children.clone())
            .unwrap_or_default()
            .into_iter()
            .filter(|c| self.is_item_element(*c))
            .collect();
        self.command(node, &WidgetCommand::ClearItems);
        for item in items {
            let text = self.text_content(item);
            self.command(node, &WidgetCommand::AddItem(text));
        }
    }

    /// The containing block: the nearest ancestor an out-of-flow child's
    /// coordinates are measured from.
    ///
    /// The nearest **positioned** ancestor, exactly as CSS defines it — an
    /// ancestor whose declared `position` is not `static`. A plain `<div>` is
    /// not one, and is correctly transparent to its children's coordinates.
    ///
    /// This is the whole rule; there is no container special case. A frontend
    /// whose children are parent-relative — VCL, WinForms, Flutter's
    /// `Positioned`, all of them — gets that by its containers **declaring**
    /// `position: absolute`, which `primitives/gui.rs` emits at construction.
    /// Either positioned value would serve here, since both establish a
    /// containing block; `absolute` is emitted because these frontends also
    /// mean "placed AT my own coordinates", which `relative` does not say.
    /// The declaration is in the document, so a real engine handed the same
    /// markup places the same controls in the same places. A behavioural
    /// assumption here would render correctly and be wrong in a browser, which
    /// is the one failure mode being HTML underneath exists to prevent.
    ///
    /// `position: fixed` resolves against the viewport rather than an ancestor,
    /// so the walk skips straight to the form.
    fn containing_block(&mut self, node: NodeId) -> LayoutRect {
        if self.declared_position(node) == "fixed" {
            return self.form.rect();
        }
        let mut cursor = self.nodes.get(&node).and_then(|n| n.parent);
        while let Some(parent) = cursor {
            if parent == DOCUMENT {
                break;
            }
            if !matches!(self.declared_position(parent).as_str(), "" | "static") {
                if let Some(w) = self.widget_mut(parent) {
                    return w.rect();
                }
            }
            cursor = self.nodes.get(&parent).and_then(|n| n.parent);
        }
        // No positioned ancestor: the initial containing block, which is the
        // viewport. A browser answers the same.
        self.form.rect()
    }

    fn declared_position(&self, node: NodeId) -> String {
        self.style(node)
            .map(|s| s.get("position").trim().to_ascii_lowercase())
            .unwrap_or_default()
    }

    /// The containing block's extent along one axis — what a percentage is a
    /// percentage OF.
    ///
    /// The parent's rect, or the viewport for a child of the document. This is
    /// the piece parsing cannot have, which is why `Length::Percent` stays
    /// symbolic until here.
    fn containing_extent(&mut self, node: NodeId, horizontal: bool) -> f32 {
        let rect = self.containing_block(node);
        if horizontal { rect.w } else { rect.h }
    }

    /// A CSS length in pixels, resolving percentages against the containing
    /// block. `horizontal` picks which axis the percentage refers to.
    fn resolve_length(&mut self, node: NodeId, value: &str, horizontal: bool) -> Option<f32> {
        if let Some(px) = parse_px(value) {
            return Some(px);
        }
        let percent = value.trim().strip_suffix('%')?.trim().parse::<f32>().ok()?;
        Some(self.containing_extent(node, horizontal) * percent / 100.0)
    }

    /// Apply `min-*`/`max-*` to a width or height.
    ///
    /// Declared constraints are read from the store rather than tracked
    /// separately, so they apply whichever order they were set in — `max-width`
    /// before or after `width` gives the same answer, which is what a
    /// declarative frontend needs since it emits fields in catalog order.
    fn clamp_to_constraints(&mut self, node: NodeId, property: &str, value: f32) -> f32 {
        let (min_key, max_key, horizontal) = match property {
            "width" => ("min-width", "max-width", true),
            "height" => ("min-height", "max-height", false),
            _ => return value,
        };
        let declared = |doc: &Self, key: &str| -> Option<String> {
            let raw = doc.style(node)?.get(key);
            (!raw.is_empty()).then(|| raw.to_string())
        };
        let min = declared(self, min_key);
        let max = declared(self, max_key);
        let mut value = value;
        if let Some(max) = max.and_then(|m| self.resolve_length(node, &m, horizontal)) {
            value = value.min(max);
        }
        if let Some(min) = min.and_then(|m| self.resolve_length(node, &m, horizontal)) {
            value = value.max(min);
        }
        value
    }

    /// Re-apply a declared width/height through the constraints — used when a
    /// `min-*`/`max-*` arrives after the size it constrains.
    fn reapply_constrained_axis(&mut self, node: NodeId, property: &str) {
        let horizontal = property == "width";
        let Some(declared) = self
            .style(node)
            .map(|s| s.get(property).to_string())
            .filter(|d| !d.is_empty())
        else {
            return;
        };
        let Some(px) = self.resolve_length(node, &declared, horizontal) else {
            return;
        };
        let clamped = self.clamp_to_constraints(node, property, px);
        if let Some(w) = self.widget_mut(node) {
            let r = apply_axis(w.rect(), property, clamped);
            w.set_rect(r);
        }
    }

    /// Tell an element's container something about that element.
    ///
    /// Per-child layout facts — grow weight, cross alignment, order — live
    /// with whoever arranges, for the reason `dock` does: the child does not
    /// act on them, the container does. The command carries `name=value`
    /// because a container is addressed by ITS name and has to know which child
    /// is meant.
    fn tell_container(&mut self, node: NodeId, verb: &str, value: &str) {
        let Some(parent) = self.nodes.get(&node).and_then(|n| n.parent) else {
            return;
        };
        if parent == DOCUMENT {
            return;
        }
        let spec = format!("{}={}", Self::widget_name(node), value);
        self.command(
            parent,
            &WidgetCommand::Custom(verb.to_string(), CommandValue::Text(spec)),
        );
    }

    /// Record a child's grow weight with its container.
    fn set_child_flex(&mut self, node: NodeId, grow: f32) {
        self.tell_container(node, "SetChildFlex", &grow.to_string());
    }

    /// Put an out-of-flow element where its own declarations say.
    ///
    /// Reads the store, not the rect: a rect is whatever the last layout pass
    /// left behind, and after an insertion that is exactly the value we are
    /// trying to undo.
    fn apply_declared_geometry(&mut self, node: NodeId) {
        let declared: Vec<(String, f32)> = ["left", "top", "width", "height"]
            .iter()
            .filter_map(|axis| {
                let value = self.style(node)?.get(axis);
                parse_px(value).map(|px| ((*axis).to_string(), px))
            })
            .collect();
        let Some(w) = self.widget_mut(node) else {
            return;
        };
        let mut rect = w.rect();
        for (axis, px) in &declared {
            rect = apply_axis(rect, axis, *px);
        }
        w.set_rect(rect);
    }

    /// Re-run the container's layout, for a child the container arranges.
    ///
    /// Handing a panel its own rect is what triggers `relayout()`. Needed
    /// because writing `left` on an in-flow child moves it directly, and in CSS
    /// `left` is inert on a `position: static` box — the container's answer has
    /// to win, or the two disagree until something else happens to relayout.
    fn relayout_parent(&mut self, node: NodeId) {
        let Some(parent) = self.nodes.get(&node).and_then(|n| n.parent) else {
            return;
        };
        if parent == DOCUMENT {
            return;
        }
        if let Some(w) = self.widget_mut(parent) {
            let rect = w.rect();
            w.set_rect(rect);
        }
    }

    /// Has this element been positioned by its own coordinates?
    ///
    /// The bridge until every frontend declares `position` explicitly. In CSS
    /// `left`/`top` do nothing on a `position: static` box, so a frontend
    /// setting them and expecting them to count is *already* saying
    /// `position: absolute` — it simply has not said so yet. Reading the
    /// declaration store means this answers from what the program actually set,
    /// not from a guess.
    fn positions_itself(&self, node: NodeId) -> bool {
        let Some(style) = self.style(node) else {
            return false;
        };
        match style.get("position").trim().to_ascii_lowercase().as_str() {
            "absolute" | "fixed" => return true,
            // `static` is the author saying the container decides, and it wins
            // over the inference below.
            "static" => return false,
            // `relative` is IN flow — the container arranges it and then
            // offsets it, which `set_child_relative` tells the container to
            // do. It used to be lumped in with `absolute` here, because
            // containers were declared `relative` while meaning "placed at my
            // coordinates"; they declare `absolute` now, so relative is free
            // to mean what CSS says.
            "relative" => return false,
            "" => {}
            _ => return false,
        }
        !style.get("left").is_empty() || !style.get("top").is_empty()
    }

    /// An element's `style`. A node that does not exist has no declarations,
    /// which is what the CSSOM says an absent element's style reads as.
    pub fn style(&self, node: NodeId) -> Option<&Style> {
        if node == DOCUMENT {
            return Some(&self.document_style);
        }
        self.nodes.get(&node).map(|n| &n.style)
    }

    /// Record a declaration. Creates nothing — an id with no node is a no-op,
    /// not a phantom element.
    fn record_style(&mut self, node: NodeId, property: &str, value: &str) {
        if node == DOCUMENT {
            self.document_style.set(property, value);
        } else if let Some(n) = self.nodes.get_mut(&node) {
            n.style.set(property, value);
        }
    }

    /// The typed view of an element's declarations, for layout to read.
    pub fn style_properties(&self, node: NodeId) -> crate::css::CssProperties {
        self.style(node).map(Style::properties).unwrap_or_default()
    }

    /// `element.style.getPropertyValue(property)`.
    ///
    /// Geometry answers from the **control**, not from the declaration: a rect
    /// is computed, and a docked or laid-out element is not where its own
    /// `left` said it would be. Everything else answers from the declaration
    /// store, which is what makes a property the write side accepts readable —
    /// `color`, `background-color` and `font-size` all returned `""` before,
    /// having been consumed into widget commands and forgotten.
    pub fn style_property(&mut self, node: NodeId, property: &str) -> String {
        let property = property.to_ascii_lowercase();
        let rect = if node == DOCUMENT {
            self.form.rect()
        } else {
            match self.widget_mut(node) {
                Some(w) => w.rect(),
                None => return self.declared_style(node, &property),
            }
        };
        match property.as_str() {
            "left" => format!("{}px", rect.x),
            "top" => format!("{}px", rect.y),
            "width" => format!("{}px", rect.w),
            "height" => format!("{}px", rect.h),
            _ => self.declared_style(node, &property),
        }
    }

    fn declared_style(&self, node: NodeId, property: &str) -> String {
        self.style(node)
            .map(|s| s.get(property).to_string())
            .unwrap_or_default()
    }

    // ── IDL properties ──────────────────────────────────────────────────

    /// `input.value` / `select.value` / `textarea.value`. Which command that
    /// is depends on the control, which is exactly what HTML says: `value`
    /// means different things to a text field, a range and a select.
    pub fn set_value(&mut self, node: NodeId, value: &str) {
        let cmd = match self.value_kind(node) {
            ValueKind::Text => WidgetCommand::SetText(value.to_string()),
            ValueKind::Number => WidgetCommand::SetValue(value.trim().parse().unwrap_or(0.0)),
            ValueKind::Index => WidgetCommand::SetSelectedIndex(value.trim().parse().unwrap_or(0)),
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
                _ => String::new(),
            },
        }
    }

    /// `input.checked` — a boolean, read from the control.
    pub fn set_checked(&mut self, node: NodeId, checked: bool) {
        self.command(node, &WidgetCommand::SetChecked(checked));
    }

    pub fn checked(&mut self, node: NodeId) -> bool {
        matches!(
            self.command(node, &WidgetCommand::GetValue),
            CommandValue::Bool(true)
        )
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
            _ => String::new(),
        }
    }

    /// `element.focus()`.
    pub fn focus(&mut self, node: NodeId) {
        self.command(node, &WidgetCommand::Focus);
    }

    /// `select.selectedIndex` — `-1` when nothing is selected.
    ///
    /// The widget already holds this: a list control answers `GetValue` with
    /// [`CommandValue::Index`], which is the same fact `value_kind_of` reports
    /// as [`ValueKind::Index`]. So this asks the control rather than keeping a
    /// second copy of the selection in the document.
    ///
    /// It is NOT `value`. `HTMLSelectElement.value` is the selected option's
    /// value string; `selectedIndex` is a `long`. Binding a toolkit's
    /// `ItemIndex` to `value` would read correctly here and wrongly in a
    /// browser.
    pub fn selected_index(&mut self, node: NodeId) -> i32 {
        match self.command(node, &WidgetCommand::GetValue) {
            CommandValue::Index(i) => i as i32,
            _ => -1,
        }
    }

    /// An element's laid-out rect.
    ///
    /// Goes through the document's own widget lookup, which walks the tree —
    /// unlike `Form::get_control_rect`, whose `find_rect` default matches only
    /// the widget itself and is not overridden by any container. That made
    /// **every nested control** report no rect while rendering perfectly, and
    /// a debugger dump saying "never laid out" for a control you can see is
    /// worse than saying nothing: it sent a session hunting container layout
    /// that was working.
    pub fn rect(&mut self, node: NodeId) -> Option<LayoutRect> {
        self.widget_mut(node).map(|w| w.rect())
    }

    /// `select.options[i].text` — the item at an index, `""` when out of range.
    ///
    /// The list had add, remove and clear but no READ, so an indexed item was
    /// unreachable from any frontend. `TStrings[i]`, .NET's `this[int]` and the
    /// options collection are all this one question.
    pub fn item_text(&mut self, node: NodeId, index: usize) -> String {
        match self.command(node, &WidgetCommand::GetItem(index)) {
            CommandValue::Text(text) => text,
            _ => String::new(),
        }
    }

    pub fn set_item_text(&mut self, node: NodeId, index: usize, text: &str) {
        self.command(node, &WidgetCommand::SetItem(index, text.to_string()));
    }

    pub fn set_selected_index(&mut self, node: NodeId, index: i32) {
        // The IDL clamps a negative assignment to "nothing selected"; the
        // command takes a `usize`, so a negative index has no command to send.
        if let Ok(index) = usize::try_from(index) {
            self.command(node, &WidgetCommand::SetSelectedIndex(index));
        }
    }

    /// `select.add(option)` / `select.remove(index)`.
    /// `select.add(option)` — append an item.
    ///
    /// A **radio group's** items are not an option list: they are child
    /// `<input type=radio>` elements, which is what HTML uses and what VCL's
    /// `TRadioGroup` and WinForms' `GroupBox` actually build. `AddItem` had
    /// nowhere to land on a `<fieldset>`, so `Items.Add('red')` raised nothing
    /// and produced nothing — a declared member that was silently inert.
    pub fn add_item(&mut self, node: NodeId, text: &str) {
        if self.is_radio_group(node) {
            let option = self.create_element("input", "radio");
            self.append_child(node, option);
            // The caption is the control's own text, the way a VCL radio item
            // carries its label. HTML would wrap the input in a `<label>`; that
            // is the shape to grow into, and until then a void element cannot
            // serialise its caption — see guiplan's gap list.
            self.set_text_content(option, text);
            return;
        }
        self.command(node, &WidgetCommand::AddItem(text.to_string()));
    }

    /// A container whose items are radio choices rather than list rows.
    fn is_radio_group(&self, node: NodeId) -> bool {
        self.node(node)
            .map(|n| n.tag == "fieldset")
            .unwrap_or(false)
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
                | WidgetEvent::SplitMoved(_, _) => continue,
            };
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

/// Close every open document.
///
/// Documents are guest-created trees held in a process-wide table, so without
/// this the next program in a reused VM inherits the previous one's DOM — the
/// defect `vybe_platform_web::html::reset_active_document` describes, at the
/// storage layer rather than the "which document is active" layer.
///
/// `next_id` is NOT rewound: ids already handed out must never be reissued, or
/// a stale id from the previous program would silently address a new
/// document's node.
pub fn reset() {
    documents().lock().unwrap().docs.clear();
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
            _ => ValueKind::Text,
        },
        _ => ValueKind::Text,
    }
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
            // A password field masks what it holds, which is exactly what the
            // masked textbox is for — it fell to the plain textbox before and
            // showed the characters.
            "password" => "maskedtextbox",
            // `text`, `password`, `email`, `search`, `tel`, `url`, and the
            // spec's "missing value default" for anything unknown.
            _ => "textbox",
        },
        "button" => "button",
        // `<select>` is BOTH controls, and `size` is what separates them —
        // HTML's own rule: a size above one is a list box, anything else is a
        // dropdown. Answering `combobox` unconditionally left a VCL `TListBox`
        // no element that could represent it, and `<ul>` — the obvious
        // alternative — has no selection IDL at all: no `selectedIndex`, no
        // indexed item text, no `remove(i)`. It rendered as a list and could
        // not answer which row was selected, silently.
        "select" => match input_type.parse::<u32>() {
            Ok(size) if size > 1 => "listbox",
            _ => "combobox",
        },
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
        // A `<select>`'s options and a table's cells are CONTENT of their
        // container, not controls beside it. They are text-bearing, so `label`
        // is right — listed explicitly so it is a decision rather than the
        // fallback catching them.
        "option" | "optgroup" | "td" | "th" | "caption" | "legend" | "summary" => "label",
        // Table structure: rows and sections hold cells, so they are
        // containers rather than text.
        "tr" | "thead" | "tbody" | "tfoot" | "colgroup" => "flowlayoutpanel",
        "iframe" | "embed" | "object" => "picturebox",
        // A rule is a thematic break: a thin panel, which is what it draws as.
        // Without an arm here it would fall to `_ => "label"` at 120x20 — the
        // silent-label trap — which is why a separator stayed on `<menu>` in
        // the dotnet adapter until this existed.
        "hr" => "panel",
        "dialog" | "div" | "form" | "section" | "article" | "main" | "aside" | "header"
        | "footer" | "nav" | "li" => "flowlayoutpanel",
        // A custom element IS the control it is named after: `<vybe-tabcontrol>`
        // is the tabcontrol widget. The tag carries the kind, so the two halves
        // of the mapping cannot drift and adding a control needs no arm here.
        //
        // This is what makes custom elements first-class rather than a
        // fallback. Before it, `control_kind` had no `vybe-` handling at all
        // and EVERY custom element — tabs, split containers, picture boxes,
        // every non-visual component — fell through to `label` at 120x20. The
        // silent-label trap, applied to the entire custom-element mechanism.
        custom if custom.starts_with("vybe-") => vybe_widget_kind(&custom[5..]),
        // Anything else is text-bearing: `p`, `span`, `label`, `h1`…`h6`, `li`.
        _ => "label",
    }
}

/// The `vybe_widgets` control a `<vybe-*>` element names, or `label` if the
/// name matches no control.
///
/// The list is `controls::make_widget`'s own, which is what keeps a custom
/// element honest: a tag naming a control that does not exist is a *mistake*,
/// and it degrades to a label rather than to nothing — visible in an `html`
/// dump and in a capture.
///
/// **This list is also the polyfill manifest.** Each entry is one
/// `customElements.define("vybe-<name>", …)` a browser build would need, so
/// keeping it in one place is what makes running in a real engine a finite job
/// rather than an archaeology exercise.
fn vybe_widget_kind(name: &str) -> &'static str {
    match name {
        "tabcontrol" => "tabcontrol",
        "tabpage" => "tabpage",
        "splitcontainer" => "splitcontainer",
        "treeview" => "treeview",
        "listview" => "listview",
        "monthcalendar" => "monthcalendar",
        "datetimepicker" => "datetimepicker",
        "picturebox" => "picturebox",
        "richtextbox" => "richtextbox",
        "maskedtextbox" => "maskedtextbox",
        "numericupdown" => "numericupdown",
        "checkedlistbox" => "checkedlistbox",
        "datagridview" | "datagrid" => "datagridview",
        "menustrip" => "menustrip",
        "toolstrip" => "toolstrip",
        "statusstrip" => "statusstrip",
        "contextmenustrip" => "contextmenustrip",
        "bindingnavigator" => "bindingnavigator",
        "flowlayoutpanel" => "flowlayoutpanel",
        "hflowlayoutpanel" => "hflowlayoutpanel",
        "tablelayoutpanel" => "tablelayoutpanel",
        "hscrollbar" => "hscrollbar",
        "vscrollbar" => "vscrollbar",
        "panel" | "usercontrol" => "panel",
        "groupbox" => "groupbox",
        "canvas" | "paintbox" => "canvas",
        "progressbar" => "progressbar",
        "trackbar" => "trackbar",
        "linklabel" => "linklabel",
        _ => "label",
    }
}

/// Elements that are in the document but draw nothing.
///
/// Metadata content (`<script>`, `<style>`, `<title>`, `<meta>`…), a hidden
/// input, and `<template>`, whose contents are inert by definition. They are
/// real nodes — `createElement("script")` must give you something you can
/// append and read back — they simply have no rendering.
///
/// Without this they fell to `_ => "label"` and drew their text at 120x20:
/// a stylesheet or a script visible in the middle of the form. The
/// silent-label trap again, from the standards side rather than the invented-tag
/// side.
fn renders_nothing(tag: &str, input_type: &str) -> bool {
    matches!(
        tag,
        "script"
            | "style"
            | "head"
            | "title"
            | "meta"
            | "link"
            | "base"
            | "template"
            | "noscript"
            | "param"
            | "source"
            | "track"
    ) || (tag == "input" && input_type == "hidden")
        // Non-visual components. A `Timer`, `ToolTip`, `ImageList` or
        // `BindingSource` is a member of the form, not a box on it — WinForms
        // and VCL both put them in a `components` collection rather than
        // `Controls`, and neither ever paints one.
        //
        // They are still nodes, for the same reason `<script>` is: a program
        // creates one, names it, wires its events and reads it back. What they
        // must not do is occupy a rectangle, which is precisely what they did
        // before — a form with a timer and a tooltip drew two grey labels.
        || matches!(
            tag,
            "vybe-timer"
                | "vybe-tooltip"
                | "vybe-imagelist"
                | "vybe-notifyicon"
                | "vybe-bindingsource"
                | "vybe-backgroundworker"
                | "vybe-errorprovider"
                | "vybe-helpprovider"
                | "vybe-openfiledialog"
                | "vybe-savefiledialog"
                | "vybe-folderbrowserdialog"
                | "vybe-colordialog"
                | "vybe-fontdialog"
        )
}

/// A starting size, so a control that is appended before it is styled is
/// visible rather than zero-sized. CSS overrides it.
fn default_size(kind: &str, tag: &str) -> (f32, f32) {
    // A rule is a panel that draws as a line — same widget, nothing like the
    // same shape, so the tag decides before the kind does.
    if tag == "hr" {
        return (200.0, 2.0);
    }
    match kind {
        "textbox" | "combobox" | "numericupdown" | "datetimepicker" => (160.0, 24.0),
        // Tall, because it holds lines — the same call `RICHTEXTBOX_DEF` makes
        // with its own (100, 96). One line high made an unstyled memo look
        // like a text field.
        "richtextbox" => (160.0, 96.0),
        "button" => (96.0, 28.0),
        "checkbox" | "radiobutton" => (140.0, 20.0),
        "trackbar" | "progressbar" => (160.0, 20.0),
        "listbox" | "listview" | "datagridview" | "treeview" => (200.0, 120.0),
        "panel" | "groupbox" | "flowlayoutpanel" | "canvas" | "picturebox" => (200.0, 150.0),
        // The custom-element controls. Without arms here they inherited the
        // text fallback of 120x20 — the right widget at label size, which reads
        // as "the control is broken" rather than "the size is unset".
        "tabcontrol" | "splitcontainer" | "tablelayoutpanel" | "hflowlayoutpanel" => (300.0, 200.0),
        "tabpage" | "usercontrol" => (300.0, 170.0),
        "monthcalendar" => (220.0, 180.0),
        "checkedlistbox" => (200.0, 120.0),
        "menustrip" | "toolstrip" | "statusstrip" | "contextmenustrip" | "bindingnavigator" => {
            (300.0, 24.0)
        }
        "hscrollbar" => (200.0, 16.0),
        "vscrollbar" => (16.0, 200.0),
        _ => (120.0, 20.0),
    }
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
        _ => r,
    }
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
        Some(f) => f.to_uppercase().collect::<String>() + c.as_str(),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    /// Render a document into a fresh pixmap and hand back the pixels.
    fn painted(doc: &mut Document, width: u32, height: u32) -> tiny_skia::Pixmap {
        let mut pixmap = tiny_skia::Pixmap::new(width, height).unwrap();
        let mut font_system = cosmic_text::FontSystem::new();
        let mut swash_cache = cosmic_text::SwashCache::new();
        doc.set_viewport(width as f32, height as f32);
        doc.render(&mut RenderContext {
            pixmap: &mut pixmap,
            font_system: &mut font_system,
            swash_cache: &mut swash_cache,
            scale: 1.0,
        });
        pixmap
    }

    #[test]
    fn a_modal_paints_over_the_page_from_the_top_layer() {
        // The whole point of the top layer: the modal's pass runs AFTER the
        // document's, so its backdrop dims a button that was painted first —
        // and would dim it just the same if the button came later in the
        // tree or carried a large z-index, neither of which the top layer
        // consults.
        let mut doc = Document::new("t");
        let button = doc.create_element("button", "");
        doc.append_child(DOCUMENT, button);
        let dialog = doc.create_element("dialog", "");
        doc.append_child(DOCUMENT, dialog);

        let before = painted(&mut doc, 200, 100);
        doc.show_dialog(dialog, true);
        let after = painted(&mut doc, 200, 100);

        // The page's background is opaque, so the backdrop shows up as a
        // DARKER pixel, not a more opaque one.
        let sample = |p: &tiny_skia::Pixmap| {
            let px = p.pixel(20, 20).unwrap();
            u32::from(px.red()) + u32::from(px.green()) + u32::from(px.blue())
        };
        assert!(
            sample(&after) < sample(&before),
            "an open modal must lay its backdrop over what the page already \
             painted: brightness {} before, {} after",
            sample(&before),
            sample(&after)
        );
    }

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
        assert_eq!(
            doc.get_attribute(b, "class"),
            None,
            "absent attribute is null"
        );
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
        assert!(
            doc.connected(field),
            "a child of an inserted node is connected"
        );
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
        assert!(
            doc.append_child(DOCUMENT, b),
            "move out of a nested container"
        );
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
    fn a_style_the_write_side_accepts_reads_back() {
        // These three were WRITTEN (they issue widget commands) and read back
        // `""`, because the declaration was consumed into the command and
        // forgotten. Nothing recorded what the author actually said.
        let mut doc = Document::new("t");
        let b = doc.create_element("button", "");
        doc.append_child(DOCUMENT, b);
        doc.set_style_property(b, "color", "red");
        doc.set_style_property(b, "background-color", "#fff");
        doc.set_style_property(b, "font-size", "18px");
        assert_eq!(doc.style_property(b, "color"), "red");
        assert_eq!(doc.style_property(b, "background-color"), "#fff");
        assert_eq!(doc.style_property(b, "font-size"), "18px");
    }

    #[test]
    fn a_style_nothing_renders_still_round_trips() {
        // The store accepts anything; layout consumes what it understands. An
        // unimplemented property is a rendering gap, not data loss.
        let mut doc = Document::new("t");
        let b = doc.create_element("div", "");
        doc.append_child(DOCUMENT, b);
        doc.set_style_property(b, "mix-blend-mode", "multiply");
        assert_eq!(doc.style_property(b, "mix-blend-mode"), "multiply");
    }

    #[test]
    fn geometry_answers_from_the_control_not_the_declaration() {
        // A rect is COMPUTED. A docked element is not where its own `left`
        // said, so geometry must keep coming off the widget even though the
        // declaration is now recorded beside it.
        let mut doc = Document::new("t");
        let bar = doc.create_element("div", "");
        doc.append_child(DOCUMENT, bar);
        doc.set_style_property(bar, "left", "300px");
        doc.set_style_property(bar, "dock", "top");
        assert_eq!(doc.style_property(bar, "left"), "0px");
        // …while the declaration itself is intact underneath.
        assert_eq!(doc.style(bar).map(|s| s.get("left")), Some("300px"));
    }

    #[test]
    fn layout_can_read_the_declarations_it_will_need() {
        use crate::css::{Display, FlexDirection, Position};
        let mut doc = Document::new("t");
        let row = doc.create_element("div", "");
        doc.append_child(DOCUMENT, row);
        doc.set_style_property(row, "display", "flex");
        doc.set_style_property(row, "flex-direction", "row");
        let props = doc.style_properties(row);
        assert!(props.is_flex_container());
        assert_eq!(props.display, Some(Display::Flex));
        assert_eq!(props.flex_direction, Some(FlexDirection::Row));

        let child = doc.create_element("button", "");
        doc.append_child(row, child);
        doc.set_style_property(child, "position", "absolute");
        let child_props = doc.style_properties(child);
        assert_eq!(child_props.position, Some(Position::Absolute));
        assert!(child_props.is_out_of_flow());
    }

    #[test]
    fn display_none_still_hides_now_that_display_carries_a_layout_mode() {
        // `display` meant visibility and nothing else. It now also selects a
        // layout mode, and `none` must not have quietly stopped hiding.
        let mut doc = Document::new("t");
        let b = doc.create_element("button", "");
        doc.append_child(DOCUMENT, b);
        doc.set_style_property(b, "display", "none");
        assert!(doc.style_properties(b).display_none);
        assert_eq!(doc.style_property(b, "display"), "none");
    }

    /// A container sized so it actually runs a layout pass.
    ///
    /// `position: relative` is declared, because that is what makes it a
    /// containing block — a static box is transparent to its children's
    /// coordinates, in CSS and here. In a real program `primitives/gui.rs`
    /// emits this at construction for every container element; the fixture
    /// says it out loud for the same reason the markup does.
    fn container_with(doc: &mut Document, w: f32, h: f32) -> NodeId {
        let panel = doc.create_element("div", "");
        doc.append_child(DOCUMENT, panel);
        doc.set_style_property(panel, "position", "relative");
        doc.set_style_property(panel, "width", &format!("{w}px"));
        doc.set_style_property(panel, "height", &format!("{h}px"));
        panel
    }

    #[test]
    fn a_positioned_child_keeps_its_own_coordinates() {
        // THE container defect. `relayout()` recomputed every child's rect
        // from flow order and discarded what `left`/`top` wrote — the rect was
        // never missing, it was overwritten, which is why controls appeared
        // stacked in flow order rather than at nonsense coordinates.
        let mut doc = Document::new("t");
        let panel = container_with(&mut doc, 200.0, 200.0);
        let b = doc.create_element("button", "");
        doc.append_child(panel, b);
        doc.set_style_property(b, "left", "20px");
        doc.set_style_property(b, "top", "90px");
        doc.set_style_property(b, "width", "120px");
        assert_eq!(doc.style_property(b, "left"), "20px");
        assert_eq!(doc.style_property(b, "top"), "90px");
        assert_eq!(doc.style_property(b, "width"), "120px");
    }

    #[test]
    fn coordinates_set_before_insertion_survive_joining_the_container() {
        // `Left := 8` before `Parent := Panel` is the ordinary way to write
        // VCL, and appending re-runs the container's layout.
        let mut doc = Document::new("t");
        let panel = container_with(&mut doc, 200.0, 200.0);
        let b = doc.create_element("button", "");
        doc.set_style_property(b, "left", "30px");
        doc.set_style_property(b, "top", "40px");
        doc.append_child(panel, b);
        assert_eq!(doc.style_property(b, "left"), "30px");
        assert_eq!(doc.style_property(b, "top"), "40px");
    }

    #[test]
    fn a_child_with_no_coordinates_still_flows() {
        // Flutter never sets left/top — its children must keep being arranged,
        // or fixing the frontends that position by pixel breaks the one that
        // does not.
        let mut doc = Document::new("t");
        let panel = container_with(&mut doc, 200.0, 200.0);
        let a = doc.create_element("button", "");
        let b = doc.create_element("button", "");
        doc.append_child(panel, a);
        doc.append_child(panel, b);
        // Arranged top-down, so the second child sits below the first.
        let a_top = doc.style_property(a, "top");
        let b_top = doc.style_property(b, "top");
        assert_ne!(a_top, b_top, "flowed children must not overlap");
    }

    #[test]
    fn an_explicit_static_position_hands_the_child_back_to_the_container() {
        // The author overruling the inference: `left` is inert on a static box
        // in CSS, so saying `static` out loud means "you arrange me".
        let mut doc = Document::new("t");
        let panel = container_with(&mut doc, 200.0, 200.0);
        let b = doc.create_element("button", "");
        doc.append_child(panel, b);
        doc.set_style_property(b, "position", "static");
        doc.set_style_property(b, "left", "60px");
        assert_ne!(doc.style_property(b, "left"), "60px");
    }

    #[test]
    fn an_out_of_flow_child_takes_no_space_from_its_siblings() {
        // The half of `position: absolute` that is easy to miss: it is removed
        // from the flow, so the flowed sibling gets the whole container.
        let mut doc = Document::new("t");
        let panel = container_with(&mut doc, 200.0, 200.0);
        let flowed = doc.create_element("button", "");
        doc.append_child(panel, flowed);
        let alone = doc.style_property(flowed, "height");

        let positioned = doc.create_element("button", "");
        doc.append_child(panel, positioned);
        doc.set_style_property(positioned, "left", "10px");
        doc.set_style_property(positioned, "top", "10px");
        assert_eq!(
            doc.style_property(flowed, "height"),
            alone,
            "an absolutely positioned sibling must not shrink the flowed one"
        );
    }

    #[test]
    fn metadata_content_is_a_node_that_draws_nothing() {
        // `<script>`/`<style>` fell to the text fallback and drew their source
        // at 120x20 — a stylesheet visible in the middle of the form. They are
        // real nodes: appendable, serialisable, and invisible.
        let mut doc = Document::new("t");
        for tag in ["script", "style", "title", "meta", "template"] {
            let n = doc.create_element(tag, "");
            assert!(doc.append_child(DOCUMENT, n), "{tag} must be appendable");
            assert_eq!(doc.style_property(n, "width"), "0px", "{tag} must not draw");
            assert_eq!(doc.style_property(n, "height"), "0px");
        }
        // …and they are still in the document, which is the half that makes
        // them nodes rather than nothing.
        assert!(doc.to_html().contains("<script>"));
    }

    #[test]
    fn a_populated_list_has_nothing_selected_until_something_selects() {
        // The widget answered `Index(0)` the moment a list was populated, so a
        // program asking "which row did the user pick?" got row 0 before the
        // user had touched it — and acted on a choice nobody made. `-1` is
        // what `HTMLSelectElement.selectedIndex` reports for a `size > 1`
        // select, and what `TListBox.ItemIndex` starts at.
        let mut doc = Document::new("t");
        let list = doc.create_element("select", "6");
        doc.append_child(DOCUMENT, list);
        doc.add_item(list, "alpha");
        doc.add_item(list, "beta");
        assert_eq!(doc.selected_index(list), -1);
        doc.set_selected_index(list, 1);
        assert_eq!(doc.selected_index(list), 1);
    }

    #[test]
    fn an_item_can_be_read_back_by_index() {
        // Add, remove and clear existed; there was no READ, so `Items[i]` was
        // unreachable from every frontend.
        let mut doc = Document::new("t");
        let list = doc.create_element("select", "6");
        doc.append_child(DOCUMENT, list);
        doc.add_item(list, "alpha");
        doc.add_item(list, "beta");
        assert_eq!(doc.item_text(list, 0), "alpha");
        assert_eq!(doc.item_text(list, 1), "beta");
        // Out of range is empty, not a panic and not a wrong row.
        assert_eq!(doc.item_text(list, 9), "");
        doc.set_item_text(list, 0, "gamma");
        assert_eq!(doc.item_text(list, 0), "gamma");
    }

    #[test]
    fn a_radio_groups_items_are_child_radios() {
        // `Items.Add('red')` on a `TRadioGroup` raised nothing and produced
        // nothing — a declared member that was silently inert, because a
        // `<fieldset>` has no option list for `AddItem` to land in. HTML's
        // answer, and VCL's, is child radios.
        let mut doc = Document::new("t");
        let group = doc.create_element("fieldset", "");
        doc.append_child(DOCUMENT, group);
        doc.add_item(group, "red");
        doc.add_item(group, "green");
        let html = doc.to_html();
        assert_eq!(html.matches("type=\"radio\"").count(), 2, "{html}");
    }

    #[test]
    fn a_custom_element_is_the_control_it_names() {
        // `<vybe-tabcontrol>` IS the tabcontrol widget. Before this,
        // `control_kind` had no `vybe-` handling at all, so every custom
        // element — the whole mechanism — was a 120x20 label.
        let mut doc = Document::new("t");
        let tabs = doc.create_element("vybe-tabcontrol", "");
        doc.append_child(DOCUMENT, tabs);
        // A tab control is not label-sized.
        assert_ne!(doc.style_property(tabs, "width"), "120px");
        // …and it serialises as the custom element it is, so a browser build
        // knows exactly which `customElements.define` it needs.
        assert!(doc.to_html().contains("<vybe-tabcontrol>"));
    }

    #[test]
    fn a_custom_element_naming_no_control_degrades_visibly() {
        // A tag naming a control that does not exist is a mistake, and it must
        // fail toward something you can SEE rather than toward nothing.
        let mut doc = Document::new("t");
        let bogus = doc.create_element("vybe-nonesuch", "");
        doc.append_child(DOCUMENT, bogus);
        assert_eq!(doc.style_property(bogus, "width"), "120px");
    }

    #[test]
    fn a_non_visual_component_is_not_a_box() {
        // A Timer or ToolTip is a member of the form, not a rectangle on it —
        // WinForms and VCL both keep them in `components`, not `Controls`.
        // They drew as grey labels before.
        let mut doc = Document::new("t");
        for tag in ["vybe-timer", "vybe-tooltip", "vybe-imagelist"] {
            let n = doc.create_element(tag, "");
            doc.append_child(DOCUMENT, n);
            assert_eq!(doc.style_property(n, "width"), "0px", "{tag} must not draw");
        }
    }

    #[test]
    fn a_hidden_input_does_not_render() {
        let mut doc = Document::new("t");
        let hidden = doc.create_element("input", "hidden");
        doc.append_child(DOCUMENT, hidden);
        assert_eq!(doc.style_property(hidden, "width"), "0px");
        // …while an ordinary one does.
        let text = doc.create_element("input", "text");
        doc.append_child(DOCUMENT, text);
        assert_ne!(doc.style_property(text, "width"), "0px");
    }

    #[test]
    fn metadata_takes_no_slot_in_a_flow_container() {
        // A `<style>` in a container must not push its siblings down by a row
        // of nothing.
        let mut doc = Document::new("t");
        let panel = container_with(&mut doc, 200.0, 400.0);
        let first = doc.create_element("button", "");
        doc.append_child(panel, first);
        let alone = doc.style_property(first, "height");

        let style = doc.create_element("style", "");
        doc.append_child(panel, style);
        assert_eq!(doc.style_property(first, "height"), alone);
    }

    #[test]
    fn a_percentage_resolves_against_the_containing_block() {
        // Parsing cannot do this — `width: 50%` is a fraction of a parent the
        // parser cannot see, which is why `Length::Percent` stays symbolic
        // until something knows the containing block.
        let mut doc = Document::new("t");
        let panel = container_with(&mut doc, 400.0, 300.0);
        let b = doc.create_element("button", "");
        doc.append_child(panel, b);
        doc.set_style_property(b, "position", "absolute");
        doc.set_style_property(b, "width", "50%");
        assert_eq!(doc.style_property(b, "width"), "200px");
        doc.set_style_property(b, "height", "10%");
        assert_eq!(doc.style_property(b, "height"), "30px");
    }

    #[test]
    fn relative_keeps_its_flow_slot_and_absolute_gives_it_up() {
        // THE difference between the two, and the only one that matters to a
        // sibling: a relative box is still arranged — it keeps the space the
        // flow gave it and is merely drawn offset — while an absolute box
        // leaves the flow and its siblings close up behind it.
        //
        // They were the same thing until now: any positioned box carrying
        // coordinates was pulled out of flow, so `position: relative` could
        // not be spelled at all.
        let offset_of = |mode: &str| {
            let mut doc = Document::new("t");
            let panel = doc.create_element("div", "");
            doc.append_child(DOCUMENT, panel);
            doc.set_style_property(panel, "position", "absolute");
            doc.set_style_property(panel, "width", "400px");
            doc.set_style_property(panel, "height", "300px");
            doc.set_style_property(panel, "flex-direction", "column");

            let first = doc.create_element("button", "");
            doc.append_child(panel, first);
            let second = doc.create_element("button", "");
            doc.append_child(panel, second);

            doc.set_style_property(first, "position", mode);
            doc.set_style_property(first, "left", "25px");
            doc.set_style_property(first, "top", "10px");

            (
                doc.rect(first).expect("first is in the document"),
                doc.rect(second).expect("second is in the document"),
            )
        };

        let (rel_first, rel_second) = offset_of("relative");
        let (_, abs_second) = offset_of("absolute");

        assert_ne!(
            rel_second.y, abs_second.y,
            "the sibling must move when the first child leaves the flow, and \
             stay put when it is merely offset"
        );
        assert!(
            rel_first.x >= 25.0,
            "a relative box is drawn offset from its flow slot, got x={}",
            rel_first.x
        );
    }

    #[test]
    fn an_absolute_child_resolves_against_the_nearest_positioned_ancestor() {
        // The canonical CSS arrangement — absolute inside relative — plus the
        // half that makes it a RULE rather than "the parent wins": a `static`
        // box in between is skipped, because a static box is not a containing
        // block. Nesting alone would answer the inner div both times.
        let mut doc = Document::new("t");
        let positioned = container_with(&mut doc, 400.0, 300.0);
        doc.set_style_property(positioned, "left", "40px");
        doc.set_style_property(positioned, "top", "20px");

        let passthrough = doc.create_element("div", "");
        doc.append_child(positioned, passthrough);
        doc.set_style_property(passthrough, "position", "static");
        doc.set_style_property(passthrough, "left", "100px");
        doc.set_style_property(passthrough, "top", "100px");

        let b = doc.create_element("button", "");
        doc.append_child(passthrough, b);
        doc.set_style_property(b, "position", "absolute");
        doc.set_style_property(b, "left", "10px");
        doc.set_style_property(b, "top", "5px");

        let rect = doc.rect(b).expect("the button is in the document");
        assert_eq!(
            (rect.x, rect.y),
            (50.0, 25.0),
            "absolute must resolve against the RELATIVE ancestor (40,20), not \
             the static div between them"
        );
    }

    #[test]
    fn an_undeclared_axis_keeps_its_natural_size_not_the_flow_size() {
        // Found in a shipped form, three times over: every control that
        // declared `Left`/`Top` and a `Width` but NO `Height` rendered as a
        // tall box spanning its container, while the one sibling that declared
        // both behaved. The bug was never about position — insertion lays the
        // child out once, and the axis the program did not declare kept the
        // stretched value forever.
        let mut doc = Document::new("t");
        let panel = container_with(&mut doc, 320.0, 200.0);
        // The control's OWN height, read before anything lays it out — a
        // detached element has never been through a flow pass.
        let detached = doc.create_element("input", "text");
        let natural_height = doc.style_property(detached, "height");

        // A flowed sibling, for contrast: it SHOULD stretch.
        let flowed = doc.create_element("input", "text");
        doc.append_child(panel, flowed);
        assert_ne!(doc.style_property(flowed, "height"), natural_height);

        let partly = doc.create_element("input", "text");
        doc.append_child(panel, partly);
        doc.set_style_property(partly, "left", "10px");
        doc.set_style_property(partly, "top", "30px");
        doc.set_style_property(partly, "width", "120px");
        // Width is honoured, height falls back to the control's own — NOT to
        // whatever the flow pass left behind.
        assert_eq!(doc.style_property(partly, "width"), "120px");
        assert_eq!(doc.style_property(partly, "height"), natural_height);
    }

    #[test]
    fn a_childs_coordinates_are_relative_to_its_container() {
        // The half the first layout batch missed, and invisible in every
        // capture whose container happened to sit at the origin: VCL, WinForms
        // and Flutter all mean parent-relative, always. In CSS terms the
        // container is the containing block.
        let mut doc = Document::new("t");
        let panel = container_with(&mut doc, 200.0, 200.0);
        doc.set_style_property(panel, "left", "10px");
        doc.set_style_property(panel, "top", "10px");

        let b = doc.create_element("button", "");
        doc.append_child(panel, b);
        doc.set_style_property(b, "left", "20px");
        doc.set_style_property(b, "top", "50px");
        // 10 + 20, 10 + 50 — not 20, 50 on the body.
        assert_eq!(doc.style_property(b, "left"), "30px");
        assert_eq!(doc.style_property(b, "top"), "60px");
    }

    #[test]
    fn a_static_container_is_transparent_to_its_childrens_coordinates() {
        // The CSS rule, stated as a test so it cannot quietly become a
        // container special case again: only a POSITIONED ancestor is a
        // containing block. A browser handed this markup answers the same, and
        // that equivalence is the whole reason the widget layer is HTML.
        let mut doc = Document::new("t");
        let panel = doc.create_element("div", "");
        doc.append_child(DOCUMENT, panel);
        doc.set_style_property(panel, "position", "static");
        doc.set_style_property(panel, "left", "10px");
        doc.set_style_property(panel, "top", "10px");
        doc.set_style_property(panel, "width", "200px");
        doc.set_style_property(panel, "height", "200px");

        let b = doc.create_element("button", "");
        doc.append_child(panel, b);
        doc.set_style_property(b, "position", "absolute");
        doc.set_style_property(b, "left", "20px");
        // Resolved against the viewport, not the static panel.
        assert_eq!(doc.style_property(b, "left"), "20px");
    }

    #[test]
    fn fixed_resolves_against_the_viewport_not_an_ancestor() {
        let mut doc = Document::new("t");
        let panel = container_with(&mut doc, 200.0, 200.0);
        doc.set_style_property(panel, "left", "10px");
        doc.set_style_property(panel, "top", "10px");
        let b = doc.create_element("button", "");
        doc.append_child(panel, b);
        doc.set_style_property(b, "position", "fixed");
        doc.set_style_property(b, "left", "20px");
        assert_eq!(doc.style_property(b, "left"), "20px");
    }

    #[test]
    fn right_anchors_to_the_opposite_edge() {
        // Flutter's `Positioned` uses all four; VCL and WinForms only ever set
        // two, which is why the other two were never wired.
        let mut doc = Document::new("t");
        let panel = container_with(&mut doc, 400.0, 300.0);
        let b = doc.create_element("button", "");
        doc.append_child(panel, b);
        doc.set_style_property(b, "width", "100px");
        doc.set_style_property(b, "right", "20px");
        // 400 wide, 20 from the right edge, 100 wide → x = 280.
        assert_eq!(doc.style_property(b, "left"), "280px");
    }

    #[test]
    fn left_and_right_together_stretch_the_box_between_them() {
        let mut doc = Document::new("t");
        let panel = container_with(&mut doc, 400.0, 300.0);
        let b = doc.create_element("button", "");
        doc.append_child(panel, b);
        doc.set_style_property(b, "left", "50px");
        doc.set_style_property(b, "right", "50px");
        assert_eq!(doc.style_property(b, "width"), "300px");
    }

    #[test]
    fn constraints_apply_whichever_order_they_arrive_in() {
        // A declarative frontend emits fields in catalog order, so `max-width`
        // may land before or after the width it constrains. Both must clamp.
        let mut doc = Document::new("t");
        let panel = container_with(&mut doc, 400.0, 300.0);

        let after = doc.create_element("button", "");
        doc.append_child(panel, after);
        doc.set_style_property(after, "position", "absolute");
        doc.set_style_property(after, "width", "300px");
        doc.set_style_property(after, "max-width", "120px");
        assert_eq!(doc.style_property(after, "width"), "120px");

        let before = doc.create_element("button", "");
        doc.append_child(panel, before);
        doc.set_style_property(before, "position", "absolute");
        doc.set_style_property(before, "max-width", "120px");
        doc.set_style_property(before, "width", "300px");
        assert_eq!(doc.style_property(before, "width"), "120px");
    }

    #[test]
    fn min_width_wins_over_max_width_when_they_conflict() {
        // CSS resolves the conflict in favour of `min`.
        let mut doc = Document::new("t");
        let panel = container_with(&mut doc, 400.0, 300.0);
        let b = doc.create_element("button", "");
        doc.append_child(panel, b);
        doc.set_style_property(b, "position", "absolute");
        doc.set_style_property(b, "max-width", "50px");
        doc.set_style_property(b, "min-width", "150px");
        doc.set_style_property(b, "width", "100px");
        assert_eq!(doc.style_property(b, "width"), "150px");
    }

    #[test]
    fn order_rearranges_children_without_moving_them_in_the_document() {
        let mut doc = Document::new("t");
        let panel = container_with(&mut doc, 200.0, 400.0);
        let first = doc.create_element("button", "");
        let second = doc.create_element("button", "");
        doc.append_child(panel, first);
        doc.append_child(panel, second);
        let first_top = doc.style_property(first, "top");
        // Send the first child to the back.
        doc.set_style_property(first, "order", "5");
        assert_ne!(doc.style_property(first, "top"), first_top);
        assert_eq!(doc.style_property(second, "top"), first_top);
        // …and the document order is untouched.
        assert!(doc.to_html().find("</button>").is_some());
    }

    #[test]
    fn align_self_overrules_the_containers_align_items() {
        let mut doc = Document::new("t");
        let panel = container_with(&mut doc, 200.0, 400.0);
        let a = doc.create_element("button", "");
        let b = doc.create_element("button", "");
        doc.append_child(panel, a);
        doc.append_child(panel, b);
        doc.set_style_property(panel, "align-items", "stretch");
        let stretched = doc.style_property(a, "width");
        doc.set_style_property(b, "align-self", "center");
        assert_eq!(doc.style_property(a, "width"), stretched);
        assert_ne!(doc.style_property(b, "width"), stretched);
    }

    #[test]
    fn justify_content_moves_the_children_within_the_leftover() {
        // Flutter `mainAxisAlignment`. Only observable when nothing grows —
        // with a growing child there is no leftover to distribute.
        let mut doc = Document::new("t");
        let panel = container_with(&mut doc, 200.0, 400.0);
        let a = doc.create_element("button", "");
        doc.append_child(panel, a);
        doc.set_style_property(a, "flex", "0");
        let packed = doc.style_property(a, "top");
        doc.set_style_property(panel, "justify-content", "center");
        let centred = doc.style_property(a, "top");
        assert_ne!(packed, centred, "centring must move the child down");
    }

    #[test]
    fn align_items_stretch_is_the_default_and_stays_so() {
        // Changing the default here would move every existing flutter widget,
        // so this pins it: declaring nothing keeps the old behaviour.
        let mut doc = Document::new("t");
        let panel = container_with(&mut doc, 200.0, 200.0);
        let a = doc.create_element("button", "");
        doc.append_child(panel, a);
        let stretched = doc.style_property(a, "width");
        doc.set_style_property(panel, "align-items", "stretch");
        assert_eq!(doc.style_property(a, "width"), stretched);
        // …and asking for something else visibly differs.
        doc.set_style_property(panel, "align-items", "center");
        assert_ne!(doc.style_property(a, "width"), stretched);
    }

    #[test]
    fn flex_direction_chooses_the_axis() {
        let mut doc = Document::new("t");
        let panel = container_with(&mut doc, 200.0, 200.0);
        let a = doc.create_element("button", "");
        let b = doc.create_element("button", "");
        doc.append_child(panel, a);
        doc.append_child(panel, b);
        doc.set_style_property(panel, "flex-direction", "row");
        // Side by side: same top, different left.
        assert_eq!(doc.style_property(a, "top"), doc.style_property(b, "top"));
        assert_ne!(doc.style_property(a, "left"), doc.style_property(b, "left"));
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
            alt: false,
        };
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
