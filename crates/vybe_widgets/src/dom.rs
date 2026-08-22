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
use std::sync::{Arc, Mutex, OnceLock, RwLock};

use crate::controls::make_widget;
use crate::css::{BoxEdges, Edges, Style};
use crate::layout::{
    Dock, LayoutRect, PanelWidget, RenderContext, WidgetCommand, WidgetEvent, find_widget_mut,
    take_widget,
};
use crate::{CommandValue, Form};

/// A node handle. `0` is the document itself — `document.body`, the form.
pub type NodeId = u64;

pub const DOCUMENT: NodeId = 0;

/// What kind of node this is — DOM `nodeType`, in the kinds a rendering tree
/// needs.
///
/// **Text is a NODE, not a property of its parent.** Without it a box's text is
/// one string that `textContent` replaces wholesale, so `a <b>B</b> c` has
/// nowhere to put the trailing ` c`: the markup is expressible, the tree is
/// not. It is the same absence that leaves `<br>` homeless and makes the
/// spec's whitespace collapsing impossible to state.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum NodeKind {
    #[default]
    Element,
    /// `Node.TEXT_NODE`. Has data, no tag, no attributes, no children — and no
    /// widget, because it is content of the box it sits in rather than a box.
    Text,
    /// `Node.COMMENT_NODE`. Data, and nothing else: no box, no line, no
    /// contribution to an ancestor's `textContent`.
    ///
    /// It exists because a parser that meets `<!-- … -->` has to put it
    /// SOMEWHERE, and the alternative is dropping it — a silent divergence
    /// between the markup handed in and the tree handed back. HTML also folds
    /// two other productions into this one kind: `<![CDATA[…]]>` and `<?…>`
    /// are both parsed as comments outside foreign content (HTML §13.2.5.42,
    /// "bogus comment state"), so the HTML side needs no separate node for
    /// either.
    Comment,
    /// `Node.CDATA_SECTION_NODE` — `<![CDATA[ … ]]>`.
    ///
    /// **A `Text` node that spells itself differently.** DOM §4.10 makes
    /// `CDATASection` a subclass of `Text`, so its data counts towards an
    /// ancestor's `textContent` exactly as a text run does; only the
    /// serialisation differs. XML-only: HTML parses the same production as a
    /// comment, which is why the HTML sink never creates one.
    CData,
    /// `Node.PROCESSING_INSTRUCTION_NODE` — `<?target data?>`.
    ///
    /// The target lives in `tag` (it IS `nodeName`) and the rest in `data`.
    /// Not character data: excluded from an ancestor's `textContent` the way a
    /// comment is. XML-only, for the same reason.
    ProcessingInstruction,
    /// `Node.DOCUMENT_FRAGMENT_NODE` — a parent with no place in the document.
    ///
    /// The point of it is what INSERTING one does: the fragment's children move
    /// into the target and the fragment itself does not (DOM §4.2.1). That is
    /// what makes building a subtree off-tree and attaching it in one step a
    /// single reflow instead of one per node, and it is why a fragment cannot
    /// be modelled as "an element nobody appended" — appending THAT would put
    /// an extra level in the tree the caller never asked for.
    ///
    /// It owns no widget, for the same reason a text node owns none: it never
    /// becomes a box.
    DocumentFragment,
}

/// Which grammar a document was built from — and therefore whether names in
/// it fold case.
///
/// **The one thing an XML document and an HTML document genuinely differ in
/// down here.** HTML tag and attribute names are ASCII case-insensitive, so
/// they are folded on the way in and everything downstream — `VOID_ELEMENTS`,
/// `control_kind`, `ua::declarations_for` — compares against lowercase
/// literals. XML is case-SENSITIVE: `<Title>` and `<title>` are two elements,
/// and folding them would break `getElementsByTagName("Title")` for every
/// caller that reads an XML document.
///
/// A property of the DOCUMENT, not an argument to each call, which is exactly
/// the distinction a browser draws between `Document` and `XMLDocument`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum DocumentKind {
    #[default]
    Html,
    Xml,
}

/// One content attribute — DOM §4.9.
///
/// **A struct rather than a map entry, because an attribute is a NODE.** The
/// spec gives it a namespace, a prefix, a local name and a qualified name, and
/// a `HashMap<String, String>` can hold exactly one of those. `setAttributeNS`
/// and `getAttributeNS` had nowhere to put or find the namespace, so every
/// namespaced attribute in an XML document collapsed onto its qualified name
/// and two attributes differing only in vocabulary became one.
///
/// `prefix` and `local_name` are derived from `name` for the same reason they
/// are on an element: one name, asked three ways, cannot drift from itself.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Attribute {
    /// `Attr.namespaceURI` — `None` for the ordinary un-namespaced case,
    /// which is every attribute in an HTML document.
    pub namespace: Option<String>,
    /// `Attr.name` — the QUALIFIED name, as written: `href`, `xlink:href`.
    pub name: String,
    pub value: String,
}

impl Attribute {
    /// `Attr.prefix` — before the colon, or `None`.
    pub fn prefix(&self) -> Option<&str> {
        self.name.split_once(':').map(|(prefix, _)| prefix)
    }

    /// `Attr.localName` — the qualified name without its prefix.
    pub fn local_name(&self) -> &str {
        match self.name.split_once(':') {
            Some((_, local)) => local,
            None => &self.name,
        }
    }
}

/// `dataset.fooBar` → `data-foo-bar` (HTML §3.2.6.6).
///
/// The two halves of the `dataset` name mapping are free functions rather than
/// methods because they depend on nothing but the string, and because getting
/// them out of step is the one way `dataset` can silently read a different
/// attribute than it writes.
fn dataset_attribute(key: &str) -> String {
    let mut name = String::from("data-");
    for ch in key.chars() {
        if ch.is_ascii_uppercase() {
            name.push('-');
            name.push(ch.to_ascii_lowercase());
        } else {
            name.push(ch);
        }
    }
    name
}

/// `data-foo-bar` → `fooBar`, or `None` for an attribute that is not `data-*`.
fn dataset_key(attribute: &str) -> Option<String> {
    let rest = attribute.strip_prefix("data-")?;
    let mut key = String::new();
    let mut upper = false;
    for ch in rest.chars() {
        if ch == '-' {
            upper = true;
            continue;
        }
        if upper {
            key.extend(ch.to_uppercase());
            upper = false;
        } else {
            key.push(ch);
        }
    }
    Some(key)
}

/// The bookkeeping a document owns and a control does not.
#[derive(Clone, Debug, Default)]
pub struct DomNode {
    pub id: NodeId,
    /// `Element` or `Text` — see [`NodeKind`].
    pub kind: NodeKind,
    /// A text node's `data`. Empty for an element.
    pub data: String,
    /// The element's QUALIFIED name — `svg`, or `xsl:template`. Folded to
    /// lowercase in an HTML document and kept verbatim in an XML one; see
    /// [`DocumentKind`].
    pub tag: String,
    /// `Element.namespaceURI` — `None` for an element created without one.
    ///
    /// The URI only. `prefix` and `localName` are the two halves of `tag` and
    /// are derived from it rather than stored, because storing them would be
    /// three fields that can disagree about one name.
    pub namespace: Option<String>,
    /// `type` at creation, which with `tag` is what the spec says decides
    /// which control an `<input>` is. Kept because `input.type` is readable.
    pub input_type: String,
    /// Attributes that no `WidgetCommand` covers — `id`, `class`, `data-*`.
    /// Everything a control does own is forwarded to it instead.
    ///
    /// An ordered LIST, which is what the spec calls it, and what lets two
    /// attributes with the same local name in different namespaces coexist.
    pub attributes: Vec<Attribute>,
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
    /// The user-agent declarations that matched this element's tag.
    ///
    /// Kept apart from `style` because a UA rule is **not** an inline style:
    /// `element.style` must not read back `font-weight: bold` for a `<strong>`
    /// nobody styled, and the serialiser must not write it into `style=""`.
    /// Merging the two — UA underneath, author on top — is the cascade.
    pub ua_style: Style,
    /// The used box-model edges, resolved from `style` on every write.
    ///
    /// The *declaration* lives in `style` and the *used value* lives here, for
    /// the same reason the CSSOM separates them: `padding: 10%` reads back as
    /// `10%` and lays out as a number, and the number depends on a containing
    /// block the declaration cannot see.
    pub box_edges: BoxEdges,
    /// **The computed style** — the cascade's answer for this element.
    ///
    /// `style` and `ua_style` are DECLARATIONS; this is the VALUE, the same
    /// split the CSSOM draws between `element.style` and `getComputedStyle`,
    /// and the same one `box_edges` already draws for the box model.
    ///
    /// It is stored rather than derived per query because a cascade *resolves*
    /// — it is not a message. Deriving it on every read meant recursing to the
    /// root each time and, worse, meant there was no computed value anywhere
    /// for anything to read: it existed as a temporary and was thrown away. An
    /// inline run's style IS its parent's resolved style, and a widget that
    /// painted from the DOM would read this; neither is possible while the
    /// answer is a temporary.
    ///
    /// Kept fresh at exactly the points that can change it — a declaration
    /// write (`record_style`), the UA layer at creation, and a move in the tree
    /// (`append_child`), since what an element inherits depends on where it is.
    pub computed: crate::css::CssProperties,
}

/// One author rule: what it selects, what it declares, and how it sorts.
///
/// `specificity` and `order` are stored rather than recomputed because they are
/// the sort key and the sort happens once per stylesheet, not once per element.
struct StyleRule {
    selector: crate::selector::ComplexSelector,
    declarations: Style,
    specificity: u32,
    /// Position in the document. The tie-break the cascade requires: two rules
    /// of equal specificity are decided by which came LAST, and without it the
    /// answer would depend on hash order.
    order: usize,
}

/// An element seen through the selector engine's eyes.
///
/// The engine matches over anything that can answer four questions
/// ([`crate::selector::Element`]), which is what lets one matcher serve the
/// live tree here and a parsed tree elsewhere without either knowing about the
/// other.
#[derive(Clone, Copy)]
struct ElementRef<'a> {
    doc: &'a Document,
    node: NodeId,
}

impl crate::selector::Element for ElementRef<'_> {
    fn tag(&self) -> String {
        // **The document root IS the body**, so it answers to `body`. A script
        // that never parsed markup has no `<body>` element — `document.body`
        // returns this node — and without this a `body { … }` rule matched
        // nothing at all in a page built by `createElement`. It computed, it
        // just had no element to land on.
        if self.node == DOCUMENT {
            return "body".to_string();
        }
        self.doc
            .node(self.node)
            .map(|n| n.tag.clone())
            .unwrap_or_default()
    }

    fn state(&self, name: &str) -> bool {
        let Some((hover, checked)) = self.doc.states.get(&self.node) else {
            return false;
        };
        match name {
            "hover" => *hover,
            "checked" => *checked,
            // Only what the parser accepts can reach here; anything else is a
            // state this tree cannot answer, and `false` is the honest answer
            // for a condition that was never observed.
            _ => false,
        }
    }

    fn attribute(&self, name: &str) -> Option<String> {
        // `style` is not in the attribute map — the declaration store IS its
        // storage, and `get_attribute` serialises it on demand. A selector has
        // to ask the same way round, or `[style]` and `[style*="color"]` would
        // silently match nothing on an element that plainly has one.
        if name == "style" {
            return self
                .doc
                .style(self.node)
                .filter(|s| !s.is_empty())
                .map(|s| s.css_text());
        }
        self.doc
            .node(self.node)
            .and_then(|n| n.attribute(name).map(str::to_string))
    }

    fn parent(&self) -> Option<Self> {
        // The document is the root of the tree but is not an element, so a
        // selector cannot match it — `html > body` has nothing to say here.
        let parent = self.doc.node(self.node)?.parent?;
        self.doc.nodes.contains_key(&parent).then_some(ElementRef {
            doc: self.doc,
            node: parent,
        })
    }

    /// The previous **element** sibling — `Element.previousElementSibling`,
    /// which is what `+` and `~` are defined over.
    ///
    /// Text and comments are siblings in the tree and not in a selector:
    /// `h1 + p` matches across the newline between them, and it would stop
    /// matching the moment the markup was indented if this counted them.
    fn previous_sibling(&self) -> Option<Self> {
        let parent = self.doc.node(self.node)?.parent?;
        let siblings = self.doc.child_nodes(parent);
        let index = siblings.iter().position(|c| *c == self.node)?;
        let previous = *siblings[..index]
            .iter()
            .rev()
            .find(|id| self.doc.is_element(**id))?;
        Some(ElementRef {
            doc: self.doc,
            node: previous,
        })
    }

    fn next_sibling(&self) -> Option<Self> {
        let parent = self.doc.node(self.node)?.parent?;
        let siblings = self.doc.child_nodes(parent);
        let index = siblings.iter().position(|c| *c == self.node)?;
        // Elements only — a text node between two boxes is not a sibling any
        // selector can see, which is the same rule `previous_sibling` applies
        // going the other way.
        let next = *siblings[index + 1..]
            .iter()
            .find(|id| self.doc.is_element(**id))?;
        Some(ElementRef {
            doc: self.doc,
            node: next,
        })
    }
}

/// One DOM event, ready for dispatch: which node, and the spec's event name.
#[derive(Clone, Debug)]
pub struct DomEvent {
    pub node: NodeId,
    /// `click`, `input`, `change`, `mouseenter`, …
    pub kind: &'static str,
    /// `preventDefault()` was called — DOM §2.3.
    prevented: bool,
    /// `stopPropagation()` was called.
    stopped: bool,
}

impl DomEvent {
    pub fn new(node: NodeId, kind: &'static str) -> Self {
        DomEvent {
            node,
            kind,
            prevented: false,
            stopped: false,
        }
    }

    /// `event.preventDefault()`.
    ///
    /// Only means anything if listeners run BEFORE the default action. A design
    /// that queues events and dispatches them a frame later cannot honour it —
    /// by then there is nothing left to prevent — which is why dispatch belongs
    /// in the browser, beside the thing it is cancelling.
    pub fn prevent_default(&mut self) {
        self.prevented = true;
    }
    pub fn default_prevented(&self) -> bool {
        self.prevented
    }
    /// `event.stopPropagation()`.
    pub fn stop_propagation(&mut self) {
        self.stopped = true;
    }
    pub fn propagation_stopped(&self) -> bool {
        self.stopped
    }
}

/// A listener's handle.
///
/// `removeEventListener` needs to unsubscribe the EXACT registration, and two
/// callbacks can be indistinguishable to the browser — it never sees what they
/// are. An id issued at registration is the only thing that identifies one.
pub type ListenerId = u32;

/// What the browser stores for a listener: an OPAQUE closure.
///
/// Deliberately not a guest value. `vybe_widgets` does not depend on the
/// runtime and must not — a browser does not know what language its page is
/// written in. A host registers by wrapping its own callback in one of these.
///
/// **The handler is HANDED the document.** It does not go and find it. A
/// listener does DOM work — `getAttribute`, `setTextContent` — and if it had to
/// re-acquire the document it would re-enter the borrow the dispatcher is
/// already holding. Passing it in is what makes dispatch synchronous, and
/// synchronous dispatch is the only kind in which `preventDefault` can cancel
/// anything.
pub type EventHandler = Box<dyn Fn(&mut DomEvent, &mut Document) + Send + Sync + 'static>;

struct Registration {
    id: ListenerId,
    node: NodeId,
    kind: String,
    handler: EventHandler,
}

#[derive(Default)]
struct EventListenersInner {
    entries: Vec<Registration>,
    next_id: ListenerId,
}

/// `EventTarget` — DOM §2.7.
///
/// Shared and interior-mutable so registering takes `&self`: a page adds a
/// listener from inside a handler, which is already holding the document.
#[derive(Clone, Default)]
pub struct EventListeners {
    inner: Arc<RwLock<EventListenersInner>>,
}

impl EventListeners {
    /// `addEventListener(type, callback)`. Returns the handle to remove it.
    pub fn add(&self, node: NodeId, kind: &str, handler: EventHandler) -> ListenerId {
        let Ok(mut inner) = self.inner.write() else {
            return 0;
        };
        inner.next_id += 1;
        let id = inner.next_id;
        inner.entries.push(Registration {
            id,
            node,
            kind: kind.to_ascii_lowercase(),
            handler,
        });
        id
    }

    /// `removeEventListener` — by the handle `add` issued.
    pub fn remove(&self, id: ListenerId) {
        if let Ok(mut inner) = self.inner.write() {
            inner.entries.retain(|e| e.id != id);
        }
    }

    /// Every listener on a node goes when the node does.
    pub fn remove_all(&self, node: NodeId) {
        if let Ok(mut inner) = self.inner.write() {
            inner.entries.retain(|e| e.node != node);
        }
    }

    pub fn is_empty(&self) -> bool {
        self.inner.read().map(|i| i.entries.is_empty()).unwrap_or(true)
    }

    /// `dispatchEvent` — run this event's listeners in registration order and
    /// answer whether the default action survives (`false` when one called
    /// `preventDefault`, matching the IDL).
    ///
    /// The table is borrowed SHARED and the document mutably, which is why the
    /// two are separate: a handler mutates the tree while the registry it came
    /// from stays readable, so a listener may add or remove listeners while
    /// running.
    /// ⚠ The read lock is HELD while handlers run, so a handler must not
    /// register or remove a listener — that takes the write lock and would
    /// deadlock. The same contract the other engine works under.
    pub fn dispatch(&self, document: &mut Document, event: &mut DomEvent) -> bool {
        // Phase 1 — which registrations match, by index, under a read borrow
        // that ends here.
        let matches: Vec<usize> = {
            let Ok(inner) = self.inner.read() else {
                return true;
            };
            inner
                .entries
                .iter()
                .enumerate()
                .filter(|(_, e)| e.node == event.node && e.kind == event.kind)
                .map(|(i, _)| i)
                .collect()
        };
        if matches.is_empty() {
            return true;
        }

        // Phase 2 — call them, handing each the document. The table is borrowed
        // shared and the document mutably, which is the whole reason they are
        // separate values.
        let Ok(inner) = self.inner.read() else {
            return true;
        };
        for index in matches {
            (inner.entries[index].handler)(event, document);
            if event.propagation_stopped() {
                break;
            }
        }
        !event.default_prevented()
    }
}

impl DomNode {
    /// An attribute by QUALIFIED name — `getAttribute`'s lookup.
    pub fn attribute(&self, name: &str) -> Option<&str> {
        self.attributes
            .iter()
            .find(|a| a.name == name)
            .map(|a| a.value.as_str())
    }

    /// An attribute by namespace and LOCAL name — `getAttributeNS`'s lookup,
    /// which is a different question: `xlink:href` and `href` share a local
    /// name and are two attributes.
    pub fn attribute_ns(&self, namespace: Option<&str>, local_name: &str) -> Option<&str> {
        self.attributes
            .iter()
            .find(|a| a.namespace.as_deref() == namespace && a.local_name() == local_name)
            .map(|a| a.value.as_str())
    }

    pub fn has_attribute(&self, name: &str) -> bool {
        self.attributes.iter().any(|a| a.name == name)
    }

    /// Set by qualified name, replacing in place so document order is stable —
    /// re-setting an attribute must not move it to the end.
    pub fn set_attribute(&mut self, namespace: Option<String>, name: &str, value: &str) {
        if let Some(existing) = self.attributes.iter_mut().find(|a| a.name == name) {
            existing.namespace = namespace;
            existing.value = value.to_string();
            return;
        }
        self.attributes.push(Attribute {
            namespace,
            name: name.to_string(),
            value: value.to_string(),
        });
    }

    pub fn remove_attribute(&mut self, name: &str) {
        self.attributes.retain(|a| a.name != name);
    }
}

/// A document and the widget tree that IS its rendering.
pub struct Document {
    /// The document element. A form's controls are the body's children, and
    /// `form.title` is `document.title` — not a copy of it.
    form: Form,
    /// **Live pseudo-class state, as the cascade last saw it** — `:hover` and
    /// `:checked` per node.
    ///
    /// A cache rather than a live query, for two reasons that happen to agree.
    /// Matching is `&self` (a selector is matched against an `ElementRef`) and
    /// asking a widget needs `&mut` — there is no immutable widget traversal,
    /// only `children_mut`. And a state that changed has to RESTYLE the
    /// element, which a query at match time could never notice: the cascade
    /// caches computed values, so `a:hover` would apply only on whatever pass
    /// happened to run next.
    ///
    /// So the state is pulled on a `&mut` sweep, compared against what is here,
    /// and a difference both updates this map and re-runs that element's
    /// cascade. One mechanism answers the selector and invalidates it.
    states: HashMap<NodeId, (bool, bool)>,
    /// Whether the stylesheet contains a selector that depends on sibling
    /// position, so a structural change has to restyle the siblings too.
    positional_rules: bool,
    nodes: HashMap<NodeId, DomNode>,
    /// Creation order, which is tree order for `getElementById`'s "first
    /// match" and for walking the document.
    order: Vec<NodeId>,
    next_id: NodeId,
    /// `EventTarget` — the listeners registered on this document's nodes.
    ///
    /// The BROWSER owns these, as a browser does. They used to live above the
    /// seam in `platforms/web`, which then had to drain a queue of "node N
    /// fired K" and match it against its own table to reconstruct a dispatch
    /// that had already happened. A page's listeners are document state, and
    /// this is where document state goes.
    listeners: EventListeners,
    /// Created but not yet inserted. The state a toolkit has no place for.
    detached: HashMap<NodeId, Box<dyn PanelWidget>>,
    /// The document element's own `style`. `DOCUMENT` is not in `nodes` — it is
    /// the form — so its declarations live beside them.
    document_style: Style,
    /// The document element's computed style — the root of every inheritance
    /// chain, and the reason `<body style="font-family: …">` reaches a control
    /// three containers down. Kept beside `document_style` for the same reason
    /// that is: `DOCUMENT` has no `DomNode` to hold it.
    document_computed: crate::css::CssProperties,
    /// **The author stylesheet** — every rule from every `<style>` in the
    /// document, in the order the cascade wants them.
    ///
    /// The third origin, and until selectors existed there was no way to have
    /// one: a rule needs a selector, so the cascade was UA and inline `style=""`
    /// with nothing between them. That is not a missing feature so much as a
    /// missing input.
    ///
    /// Held sorted — specificity, then source order — so applying it is a
    /// forward pass with no comparison per element.
    /// See [`DocumentKind`] — decides whether names fold case.
    kind: DocumentKind,
    stylesheet: Vec<StyleRule>,
    /// **The document's own child list.**
    ///
    /// Every other node keeps its children in its `DomNode`; the document has
    /// none, so before this its children were DERIVED by filtering creation
    /// order for nodes whose parent was `DOCUMENT`. That answers "who are they"
    /// and cannot answer "in what order" — creation order IS document order
    /// under a derivation, so `insertBefore` at the root was not expressible.
    /// It was also fragile: an attribute write on the document created an empty
    /// `DomNode` for it, and the derivation stopped being reached at all.
    ///
    /// A real list, in the one place the document's other bookkeeping already
    /// lives.
    root_children: Vec<NodeId>,
    /// Which boxes currently hold inline runs.
    ///
    /// Kept so that removing the last inline child still sends an empty list —
    /// otherwise the box would keep painting the runs it was last told, which
    /// is the stale-copy failure this whole migration exists to remove.
    inline_content: HashMap<NodeId, bool>,
    /// The answer for a node that does not exist.
    ///
    /// A missing element has no computed style, and the CSSOM's answer for one
    /// is "everything initial" rather than an error. Held as a field so
    /// `computed_style` can hand out a reference without every caller taking an
    /// `Option` it would only `unwrap_or_default()`.
    empty_computed: crate::css::CssProperties,
    /// THE TOP LAYER — HTML's own name for it, and its own ordered set.
    ///
    /// `showModal()` puts a dialog here; `close()` takes it out. Membership is
    /// what makes a modal paint above everything, and it is deliberately NOT
    /// a `z-index`: the top layer outranks the whole cascade, which is why a
    /// modal beats a `z-index: 9999` sibling in a browser. Modelling it as a
    /// separate paint pass is what makes that true here rather than
    /// approximately true.
    top_layer: Vec<NodeId>,
    /// `document.body`, as an ELEMENT. See [`Document::body`].
    body: Option<NodeId>,
}

impl Document {
    /// An XML document — names keep their case. See [`DocumentKind`].
    pub fn new_xml(title: &str) -> Self {
        let mut document = Self::new(title);
        document.kind = DocumentKind::Xml;
        document
    }

    /// Which grammar this document was built from.
    pub fn kind(&self) -> DocumentKind {
        self.kind
    }

    /// Fold a tag or attribute name the way THIS document does.
    ///
    /// One function, so the four places that take a name — creating an
    /// element, writing an attribute, reading one back, looking up by tag —
    /// cannot disagree about whether this document folds. They did not
    /// disagree before only because every one of them folded unconditionally.
    fn fold_name(&self, name: &str) -> String {
        match self.kind {
            DocumentKind::Html => name.to_ascii_lowercase(),
            DocumentKind::Xml => name.to_string(),
        }
    }

    pub fn new(title: &str) -> Self {
        let mut form = Form::new(title);
        // The initial viewport. A document always has one — hit testing and
        // percentage layout both need a containing block, and `window.open`
        // overrides it from the `features` string.
        <Form as PanelWidget>::set_rect(&mut form, LayoutRect::new(0.0, 0.0, 800.0, 600.0));
        let mut document = Document {
            form,
            states: HashMap::new(),
            positional_rules: false,
            nodes: HashMap::new(),
            order: Vec::new(),
            next_id: 0,
            listeners: EventListeners::default(),
            detached: HashMap::new(),
            document_style: Style::new(),
            document_computed: crate::css::CssProperties::default(),
            kind: DocumentKind::Html,
            stylesheet: Vec::new(),
            root_children: Vec::new(),
            inline_content: HashMap::new(),
            empty_computed: crate::css::CssProperties::default(),
            top_layer: Vec::new(),
            body: None,
        };
        // **`<html>`, `<head>` and `<body>` always exist** — HTML tree
        // construction inserts all three when the markup omits them, which is
        // why `<p>hi` is a complete document. A page that never writes them is
        // the common case, not the degenerate one.
        //
        // They are created HERE, rather than only when a parser sees the tags,
        // because `document.body` answered the DOCUMENT node itself when no
        // body element existed. That conflation is what made the root a special
        // case: the document is not a box, so it could not run a formatting
        // context, and everything appended to `body` had to be laid out by hand
        // (`relayout_body`) instead of by the flow every other block box uses.
        // With a real element there, `body` is an ordinary block container and
        // the root stops being different from a `<div>`.
        let html = document.create_element_typed("html", "");
        document.append_child(DOCUMENT, html);
        let head = document.create_element_typed("head", "");
        document.append_child(html, head);
        let body = document.create_element_typed("body", "");
        document.append_child(html, body);
        document.body = Some(body);
        // **The root boxes start at zero inset.** `FlowLayoutPanel` gives
        // itself 4px of padding as a widget affordance — sensible for a panel
        // someone drops on a form, wrong for the initial containing block,
        // where it silently shrinks the content width and every percentage
        // resolved against it (a block filled 792 of an 800px viewport).
        //
        // The browser's own `body { margin: 8px }` is a UA DECLARATION, not a
        // widget default, and belongs in `ua.rs` where it can be overridden by
        // a stylesheet the way it is in a browser.
        for root in [html, head, body] {
            document.command(
                root,
                &WidgetCommand::Custom("SetPadding".into(), CommandValue::Number(0.0)),
            );
        }
        document.fit_body_to_viewport();
        document
    }

    /// `document.body` — the element, never the document.
    ///
    /// Always `Some` for an HTML document: [`Document::new`] runs the tree
    /// construction that guarantees it. Returned as an `Option` only because an
    /// XML document has no body, and answering the root there would recreate
    /// exactly the conflation this exists to remove.
    pub fn body(&self) -> Option<NodeId> {
        self.body
    }

    /// Where a node addressed to the DOCUMENT actually goes: the **body**.
    ///
    /// Tree construction puts content in the body when the markup omits the
    /// tags, and a browser refuses outright to give a Document a second element
    /// child — `document.appendChild(<p>)` is a `HierarchyRequestError`. So a
    /// caller that says "the document" means the body, and this is the one
    /// place that says so.
    ///
    /// `<html>`, `<head>` and `<body>` are the exception: they ARE the
    /// document's structure, so a parser that spells them out is obeyed.
    ///
    /// Shared by every mutation rather than applied at `insert_before` alone —
    /// `remove_child` and `replace_child` unlink from the parent they are
    /// GIVEN, so a redirect on the way in and not on the way out would leave a
    /// node linked into the body and unlinked from the document, which is a
    /// tree that disagrees with itself.
    fn content_parent(&self, parent: NodeId, child: NodeId) -> NodeId {
        if parent != DOCUMENT {
            return parent;
        }
        let structural = self
            .node(child)
            .map(|n| matches!(n.tag.as_str(), "html" | "head" | "body"))
            .unwrap_or(false);
        match (structural, self.body) {
            (false, Some(body)) => body,
            _ => parent,
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
        self.fit_body_to_viewport();
    }

    /// `<html>` and `<body>` FILL the viewport.
    ///
    /// The initial containing block is the viewport, and the root boxes take
    /// their width from it — so every percentage, every `right`/`bottom` anchor
    /// and every block that fills its container resolves against the window and
    /// not against a panel's default size. Without this the body kept
    /// `FlowLayoutPanel`'s own default, and content that used to resolve
    /// against the viewport (because it hung off the document directly)
    /// silently started resolving against a 300x200 box.
    fn fit_body_to_viewport(&mut self) {
        let viewport = self.form.rect();
        let roots: Vec<NodeId> = self
            .body
            .into_iter()
            .chain(self.body.and_then(|b| self.node(b).and_then(|n| n.parent)))
            .filter(|node| *node != DOCUMENT)
            .collect();
        for node in roots {
            if let Some(widget) = self.widget_mut(node) {
                widget.set_rect(LayoutRect::new(0.0, 0.0, viewport.w, viewport.h));
            }
        }
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
        dom_node.set_attribute(None, "open", "");
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
            dom_node.remove_attribute("open");
        }
        self.command(node, &WidgetCommand::SetVisible(false));
        self.top_layer.retain(|open| *open != node);
    }

    /// `dialog.open` — read straight off the reflected attribute.
    pub fn dialog_open(&self, node: NodeId) -> bool {
        self.nodes
            .get(&node)
            .is_some_and(|n| n.has_attribute("open"))
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
    /// `document.createElement(tagName)` — DOM §4.5.
    ///
    /// ONE argument, as the IDL has it. An `<input>` starts as its missing-value
    /// default (`text`) and becomes a checkbox when `type` is set, exactly as it
    /// does in a browser — `set_attribute("type", …)` already rebuilds the
    /// control, and its own comment records what it cost to learn that ("every
    /// `<input>` a script built rendered as a text box, because the type always
    /// arrived after the element did").
    ///
    /// This used to take the type as a second parameter, which saved that one
    /// rebuild but put an engine detail — that this browser realises a control
    /// eagerly — into the shared surface, where the other engine has no such
    /// need and no such argument.
    pub fn create_element(&mut self, tag: &str) -> NodeId {
        self.create_element_typed(tag, "")
    }

    /// Create an element WITH its input type already known, skipping the
    /// rebuild that setting `type` afterwards would cause.
    ///
    /// Not the web API — an internal shortcut for callers that already know
    /// both, kept because this browser builds the control at creation time.
    pub fn create_element_typed(&mut self, tag: &str, input_type: &str) -> NodeId {
        let tag = self.fold_name(tag.trim());
        // The disambiguator is a VALUE, not a name — `<input type=CheckBox>`
        // names the same control as `checkbox` in HTML, and `control_kind`
        // compares it against lowercase literals. It folds with the tag in an
        // HTML document and, like every other value, is left alone in XML.
        let input_type = self.fold_name(input_type.trim());
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
        // The user-agent stylesheet, applied before anything else can write —
        // which is exactly what makes the cascade UA < author without a
        // priority flag beside each declaration. A browser does the same thing
        // in the same order.
        let ua: Vec<_> = {
            let tag = &self.nodes[&id].tag;
            crate::ua::declarations_for(tag).collect()
        };
        // Every UA declaration is RECORDED first, then the cascade is run once,
        // then they are applied. Interleaving the three meant each declaration
        // was applied against a cascade that had only seen the ones before it —
        // fine while `style_properties` re-derived from the declaration stores,
        // and wrong the moment the computed style became a stored value.
        for (property, value) in &ua {
            if let Some(n) = self.nodes.get_mut(&id) {
                n.ua_style.set(property, value);
            }
        }
        self.resolve_computed(id);
        for (property, value) in ua {
            self.resolve_box_edges(id);
            self.apply_style_property(id, property, value);
        }
        id
    }

    /// `document.createElementNS(namespace, qualifiedName)` — DOM §4.5.
    ///
    /// The namespace is recorded; the qualified name is the tag. Nothing else
    /// about the element changes, which is the point: a namespaced element is
    /// an ordinary element that also knows which vocabulary it came from.
    pub fn create_element_ns(
        &mut self,
        namespace: &str,
        qualified_name: &str,
        input_type: &str,
    ) -> NodeId {
        let node = self.create_element_typed(qualified_name, input_type);
        if let Some(n) = self.nodes.get_mut(&node) {
            // An empty namespace is `null`, per spec — not the empty string.
            n.namespace = (!namespace.is_empty()).then(|| namespace.to_string());
        }
        node
    }

    /// `Element.namespaceURI`.
    pub fn namespace_uri(&self, node: NodeId) -> Option<String> {
        self.node(node)?.namespace.clone()
    }

    /// `Element.prefix` — the part of the qualified name before the colon, or
    /// `None` when there is none.
    pub fn prefix(&self, node: NodeId) -> Option<String> {
        let tag = self.node(node)?.tag.clone();
        tag.split_once(':').map(|(prefix, _)| prefix.to_string())
    }

    /// `Element.localName` — the qualified name without its prefix.
    ///
    /// Derived rather than stored for the same reason `prefix` is: one name,
    /// asked two ways, cannot drift from itself.
    pub fn local_name(&self, node: NodeId) -> String {
        let Some(node) = self.node(node) else {
            return String::new();
        };
        match node.tag.split_once(':') {
            Some((_, local)) => local.to_string(),
            None => node.tag.clone(),
        }
    }

    /// `document.createTextNode(data)`.
    ///
    /// Creates a node and no widget: a text node is CONTENT of the box it is
    /// appended to, drawn as a run of that box's line, so it has no rect and
    /// nothing to lay out. It is still a full node — it has an id, a parent and
    /// a position among its siblings, and that position is what makes
    /// `a <b>B</b> c` expressible at all.
    pub fn create_text_node(&mut self, data: &str) -> NodeId {
        self.next_id += 1;
        let id = self.next_id;
        self.nodes.insert(
            id,
            DomNode {
                id,
                kind: NodeKind::Text,
                data: data.to_string(),
                ..DomNode::default()
            },
        );
        self.order.push(id);
        id
    }

    /// `document.createComment(data)`.
    ///
    /// Same storage as a text node and the opposite rendering: a text node is
    /// content of its parent's line, a comment is content of nothing at all.
    pub fn create_comment(&mut self, data: &str) -> NodeId {
        self.next_id += 1;
        let id = self.next_id;
        self.nodes.insert(
            id,
            DomNode {
                id,
                kind: NodeKind::Comment,
                data: data.to_string(),
                ..DomNode::default()
            },
        );
        self.order.push(id);
        id
    }

    /// `document.createDocumentFragment()` — a detached parent to build in.
    ///
    /// No widget and no tag: it is never a box. What it is FOR is the moment
    /// it gets appended, when its children move across and it does not — see
    /// [`append_child`](Self::append_child).
    pub fn create_document_fragment(&mut self) -> NodeId {
        self.next_id += 1;
        let id = self.next_id;
        self.nodes.insert(
            id,
            DomNode {
                id,
                kind: NodeKind::DocumentFragment,
                tag: "#document-fragment".to_string(),
                ..DomNode::default()
            },
        );
        self.order.push(id);
        id
    }

    /// Whether this node is a `DocumentFragment` — the question
    /// [`append_child`](Self::append_child) and
    /// [`insert_before`](Self::insert_before) ask before deciding what
    /// inserting it means.
    pub fn is_document_fragment(&self, node: NodeId) -> bool {
        self.node(node).map(|n| n.kind) == Some(NodeKind::DocumentFragment)
    }

    /// `node.nodeType` — the WHATWG number (DOM §4.4).
    ///
    /// `9` for the document, which has no `DomNode` of its own; the kinds
    /// carry the rest. The numbers are the spec's and are not contiguous —
    /// `2` is `Attr` and `5`/`6` are gone from the standard entirely — so this
    /// is a table rather than a cast of the enum's discriminant.
    pub fn node_type(&self, node: NodeId) -> u16 {
        if node == DOCUMENT {
            return 9;
        }
        match self.node(node).map(|n| n.kind) {
            Some(NodeKind::Element) => 1,
            Some(NodeKind::Text) => 3,
            Some(NodeKind::CData) => 4,
            Some(NodeKind::ProcessingInstruction) => 7,
            Some(NodeKind::Comment) => 8,
            Some(NodeKind::DocumentFragment) => 11,
            None => 0,
        }
    }

    /// `node.nodeName` — the tag for an element, the target for a processing
    /// instruction, and the spec's fixed names for everything else.
    pub fn node_name(&self, node: NodeId) -> String {
        if node == DOCUMENT {
            return "#document".to_string();
        }
        match self.node(node).map(|n| (n.kind, n.tag.clone())) {
            Some((NodeKind::Element, tag)) => tag,
            Some((NodeKind::ProcessingInstruction, target)) => target,
            Some((NodeKind::Text, _)) => "#text".to_string(),
            Some((NodeKind::CData, _)) => "#cdata-section".to_string(),
            Some((NodeKind::Comment, _)) => "#comment".to_string(),
            Some((NodeKind::DocumentFragment, _)) => "#document-fragment".to_string(),
            None => String::new(),
        }
    }

    /// `node.nodeValue` — the data for every node that has some, and `None`
    /// for an element or the document. Distinct from `""`, which is what an
    /// EMPTY comment answers.
    pub fn node_value(&self, node: NodeId) -> Option<String> {
        if node == DOCUMENT || self.is_element(node) {
            return None;
        }
        self.node(node).map(|n| n.data.clone())
    }

    /// `node.parentNode` — `None` for a detached node and for the document.
    pub fn parent_node(&self, node: NodeId) -> Option<NodeId> {
        self.node(node)?.parent
    }

    /// Is this node a text node?
    pub fn is_text_node(&self, node: NodeId) -> bool {
        self.node(node)
            .map(|n| n.kind == NodeKind::Text)
            .unwrap_or(false)
    }

    /// `document.createCDATASection(data)`.
    pub fn create_cdata_section(&mut self, data: &str) -> NodeId {
        self.create_data_node(NodeKind::CData, "", data)
    }

    /// `document.createProcessingInstruction(target, data)`.
    ///
    /// The target is the node's NAME, so it lands in `tag` — the same field
    /// `nodeName` reads for an element. A PI is not an element, but it is the
    /// only other node kind that HAS a name, and giving it a second field to
    /// live in would mean two places to ask.
    pub fn create_processing_instruction(&mut self, target: &str, data: &str) -> NodeId {
        self.create_data_node(NodeKind::ProcessingInstruction, target, data)
    }

    /// The shared half of every non-element factory.
    fn create_data_node(&mut self, kind: NodeKind, tag: &str, data: &str) -> NodeId {
        self.next_id += 1;
        let id = self.next_id;
        self.nodes.insert(
            id,
            DomNode {
                id,
                kind,
                tag: tag.to_string(),
                data: data.to_string(),
                ..DomNode::default()
            },
        );
        self.order.push(id);
        id
    }

    /// Does this node's data count towards an ancestor's `textContent`?
    ///
    /// Text and CDATA yes — DOM §4.10 makes `CDATASection` a `Text` — comments
    /// and processing instructions no. Their own `textContent` is still their
    /// data; the two rules disagree on purpose.
    pub fn is_character_data(&self, node: NodeId) -> bool {
        self.node(node)
            .map(|n| matches!(n.kind, NodeKind::Text | NodeKind::CData))
            .unwrap_or(false)
    }

    /// Is this node one that carries data and never a box — a comment, a CDATA
    /// section or a processing instruction?
    ///
    /// The three that attach ANYWHERE: unlike a text node they place no
    /// requirement on the parent, because they contribute no line.
    fn is_detached_data(&self, node: NodeId) -> bool {
        self.node(node)
            .map(|n| {
                matches!(
                    n.kind,
                    NodeKind::Comment | NodeKind::CData | NodeKind::ProcessingInstruction
                )
            })
            .unwrap_or(false)
    }

    /// Is this node a comment?
    pub fn is_comment_node(&self, node: NodeId) -> bool {
        self.node(node)
            .map(|n| n.kind == NodeKind::Comment)
            .unwrap_or(false)
    }

    /// Is this node an element?
    ///
    /// The question every enumeration has to ask now that the tree holds more
    /// than elements. `querySelectorAll` matches ELEMENTS — `*` is not a text
    /// node — and so does every combinator that walks siblings.
    pub fn is_element(&self, node: NodeId) -> bool {
        self.node(node)
            .map(|n| n.kind == NodeKind::Element)
            .unwrap_or(false)
    }

    /// A text node's `data`, or the empty string for an element.
    pub fn text_data(&self, node: NodeId) -> String {
        self.node(node)
            .filter(|n| n.kind == NodeKind::Text)
            .map(|n| n.data.clone())
            .unwrap_or_default()
    }

    /// Replace a text node's data.
    pub fn set_text_data(&mut self, node: NodeId, data: &str) {
        let parent = match self.nodes.get_mut(&node) {
            Some(n) if n.kind == NodeKind::Text => {
                n.data = data.to_string();
                n.parent
            }
            _ => return,
        };
        if let Some(parent) = parent {
            self.rebuild_inline_content(parent);
        }
    }

    /// The `<canvas>` element's drawing surface.
    ///
    /// **The hop that was missing between two trees.** `web:canvas` resolved
    /// its target through `GuiState.form.controls` while `web:dom` resolved
    /// through this document, so a `<canvas>` created with `createElement`
    /// received paint ops addressed to a different canvas — two surfaces, and
    /// nothing rendered. Finding the NODE was never the gap: `getElementById`
    /// has been here and exposed as `web:html:getElementById` since the
    /// conversion. Reaching the widget from it was.
    ///
    /// `as_any_mut` exists on `PanelWidget` for exactly this and says so; only
    /// `Canvas` overrides it.
    /// The canvas ELEMENT's widget.
    ///
    /// Not public: a page does not reach an element's rendering object, it asks
    /// for a CONTEXT. `get_context_2d` below is that, and this is the hop it
    /// takes to get there.
    fn canvas_mut(&mut self, node: NodeId) -> Option<&mut crate::canvas_widget::Canvas> {
        self.widget_mut(node)?
            .as_any_mut()?
            .downcast_mut::<crate::canvas_widget::Canvas>()
    }

    /// `canvas.getContext("2d")` — the drawing surface, or `None` when the node
    /// is not a `<canvas>`.
    ///
    /// One call, not two. Callers used to write
    /// `canvas_mut(node)?.canvas_mut()`: the element's widget, then the surface
    /// inside it. That second hop is an implementation detail of how a canvas
    /// is stored, and the IDL has no name for it — `getContext` goes straight
    /// from the element to the context.
    pub fn get_context_2d(&mut self, node: NodeId) -> Option<&mut crate::canvas::RecordingCanvas> {
        Some(self.canvas_mut(node)?.canvas_mut())
    }

    /// `context.measureText(text).width` — the one canvas call that ASKS.
    ///
    /// Answers the advance width of `text` in the font currently in effect on
    /// that `<canvas>`, in CSS pixels. `None` when the node is not a canvas,
    /// which is the whole reason it is an `Option`: a caller has to be able to
    /// tell "there is no surface" from "the text is zero wide", and a
    /// plausible `0.0` would make an adapter lay out against a lie.
    ///
    /// Shaping is done with the SAME attributes `fillText` will draw with —
    /// measuring in the regular face and drawing in bold is how text overruns
    /// the box it was sized for.
    pub fn measure_text(&mut self, node: NodeId, text: &str) -> Option<f32> {
        let font = self.get_context_2d(node)?.current_font();
        let spec = crate::ide_text::FontSpec {
            family: font.family.clone(),
            size: font.size,
            weight: match font.weight {
                crate::canvas::FontWeight::Bold => 700,
                crate::canvas::FontWeight::Normal => 400,
            },
            italic: matches!(font.style, crate::canvas::FontStyle::Italic),
            underline: false,
            line_through: false,
            line_height: None,
            // Canvas text has no CSS box, so no `letter-spacing` reaches it —
            // `ctx.letterSpacing` is a separate canvas attribute this does not
            // implement yet.
            letter_spacing: 0.0,
        };
        // Scale 1.0: the canvas coordinate system is CSS pixels, and the
        // device scale is applied when the recording is replayed onto a
        // pixmap. Measuring at the device scale would answer in the wrong
        // unit on any display that is not 1x.
        Some(crate::ide_text::with_font_system(|fonts| {
            crate::ide_text::measure_text_spec(fonts, text, &spec, 1.0)
        }))
    }

    /// Resolve a control NAME the way a host bridge is handed one.
    ///
    /// `getContext("myCanvas")` passes a name, not an id — SDL, .NET's
    /// `CreateGraphics` and Flutter's canvas bridge all do. So the three
    /// identities an element can be known by are tried in the order a caller
    /// means them:
    ///
    /// 1. the `name` content attribute — what `Control.Name := 'x'` sets;
    /// 2. the `id` attribute, since `getContext` also accepts an id;
    /// 3. this document's internal widget name (`n7`), which is what a handle
    ///    round-tripped through the toolkit carries.
    ///
    /// Matching is case-insensitive because the host bridge lower-cases a
    /// target name on the way in.
    pub fn element_by_control_name(&self, name: &str) -> Option<NodeId> {
        let wanted = name.trim();
        if wanted.is_empty() {
            return None;
        }
        let by_attribute = |attribute: &str| {
            self.order.iter().copied().find(|id| {
                self.nodes
                    .get(id)
                    .and_then(|n| n.attribute(attribute))
                    .map(|v| v.eq_ignore_ascii_case(wanted))
                    .unwrap_or(false)
            })
        };
        by_attribute("name")
            .or_else(|| by_attribute("id"))
            .or_else(|| {
                self.order
                    .iter()
                    .copied()
                    .find(|id| Self::widget_name(*id).eq_ignore_ascii_case(wanted))
            })
    }

    /// `document.getElementById(elementId)` — the `id` ATTRIBUTE, first match
    /// in tree order.
    pub fn get_element_by_id(&self, element_id: &str) -> Option<NodeId> {
        self.order.iter().copied().find(|id| {
            self.nodes
                .get(id)
                .and_then(|n| n.attribute("id"))
                .map(|v| v == element_id)
                .unwrap_or(false)
        })
    }

    /// `document.querySelectorAll(selectors)` — the real thing.
    ///
    /// **In tree order**, which the spec requires and which is why this walks
    /// `self.order` rather than the selector's own matches: a caller indexing
    /// `[0]` is asking for the FIRST in the document, not the first the matcher
    /// happened to reach.
    ///
    /// An invalid or unsupported selector yields nothing rather than
    /// everything. The spec throws `SyntaxError`; refusing to match is the
    /// closest answer available without an exception channel here, and it fails
    /// in the safe direction — `querySelectorAll(":hover")` returning every
    /// element would be a silent catastrophe at a call site expecting a few.
    pub fn query_selector_all(&self, selectors: &str) -> Vec<NodeId> {
        let Some(parsed) = crate::selector::parse_selector_list(selectors) else {
            return Vec::new();
        };
        self.order
            .iter()
            .copied()
            // Selectors match ELEMENTS (Selectors §3), so a text node and a
            // comment are not candidates — `*` in particular would otherwise
            // return every run of whitespace in the document.
            .filter(|id| self.is_element(*id))
            .filter(|id| {
                let element = ElementRef {
                    doc: self,
                    node: *id,
                };
                parsed
                    .iter()
                    .any(|selector| crate::selector::matches(&element, selector))
            })
            .collect()
    }

    /// `document.querySelector(selectors)` — the first match in tree order.
    pub fn query_selector(&self, selectors: &str) -> Option<NodeId> {
        self.query_selector_all(selectors).into_iter().next()
    }

    /// Every element with this tag name, in tree order.
    ///
    /// `getElementsByTagName`, not `querySelectorAll` — kept separate because
    /// it is a different API with a different argument: a tag NAME, not a
    /// selector, so `getElementsByTagName("*")` and `querySelectorAll("*")`
    /// agreeing is a coincidence rather than a shared implementation.
    pub fn get_elements_by_tag_name(&self, tag: &str) -> Vec<NodeId> {
        let tag = self.fold_name(tag);
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
                let attr = self.nodes.get(id)?.attribute("id")?;
                Some((*id, attr.to_string()))
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

    /// Whether this element paints a box of its own — a declared background.
    /// Asked of the WIDGET, so it answers what will actually be drawn rather
    /// than what the cascade holds.
    pub fn paints_own_box(&mut self, node: NodeId) -> bool {
        matches!(
            self.command(node, &WidgetCommand::Custom("PaintsBackground".into(), CommandValue::None)),
            CommandValue::Bool(true)
        )
    }

    /// Whether this element paints a border of its own.
    pub fn paints_own_border(&mut self, node: NodeId) -> bool {
        matches!(
            self.command(node, &WidgetCommand::Custom("PaintsBorder".into(), CommandValue::None)),
            CommandValue::Bool(true)
        )
    }

    /// The page's own ground — what `body { background: … }` paints.
    pub fn page_background(&self) -> (u8, u8, u8, u8) {
        self.form.background
    }

    // ── Node ────────────────────────────────────────────────────────────

    /// The unparented subtree a node currently sits in, if any. Walking up
    /// `parent` is exact — no probing — and it is what lets a node be styled
    /// or moved while its whole subtree is still outside the document.
    fn detached_root(&self, node: NodeId) -> Option<NodeId> {
        // An element whose OWN control is detached is its own root. That is the
        // inline case and it is a genuinely new state: a node with a parent in
        // the document and no widget in the tree, because it is a run of its
        // parent's text rather than a box. Walking up from it found the parent,
        // decided the node was "in the document", and sent every command to a
        // widget that was never inserted — so `set_text_content` on a `<strong>`
        // silently did nothing.
        if self.detached.contains_key(&node) {
            return Some(node);
        }
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

    /// Rebuild a node's control from what the node says it is **now**.
    ///
    /// [`Document::create_element`] decides the kind once, from an `input_type`
    /// only a caller who already knows the type can pass. The DOM's
    /// `createElement` has no such argument: `document.createElement("input")`
    /// followed by `setAttribute("type", "checkbox")` is how every script
    /// builds a control, and nothing here listened — so the element kept the
    /// text box it was born as and EVERY scripted input was a text field,
    /// whatever it claimed to be. HTML §4.10.5.3 makes a type change a change
    /// to the control itself, not to a label on it.
    ///
    /// The node, its id and its widget NAME all survive. Lookups go through
    /// `widget_name(id)`, so keeping the name is what keeps every reference —
    /// event wiring included — resolving to the control that replaced the old
    /// one.
    fn realize_control(&mut self, node: NodeId) {
        let Some(n) = self.node(node) else {
            return;
        };
        let (tag, input_type) = (n.tag.clone(), n.input_type.clone());
        let parent = n.parent;
        // Read across the swap, not after it: the old control is the only
        // thing holding these and it is about to be dropped.
        let text = self.text_content(node);
        let previous = self.get_bounding_client_rect(node);

        let kind = control_kind(&tag, &input_type);
        let (w, h) = default_size(kind, &tag);
        let mut widget = make_widget(kind, &Self::widget_name(node), &text, w, h);
        let (w, h) = if renders_nothing(&tag, &input_type) {
            (0.0, 0.0)
        } else {
            (w, h)
        };
        // The control keeps its POSITION and takes the new kind's size — a
        // swatch is not a text box's width, and moving it to the origin would
        // be a layout change nobody asked for. A control that was never laid
        // out has no position to keep.
        let origin = previous.unwrap_or(LayoutRect::zero());
        widget.set_rect(LayoutRect::new(origin.x, origin.y, w, h));
        drop(self.extract_widget(node));
        self.detached.insert(node, widget);

        // Put it back where the old one stood. Re-entering `insert_before` is
        // what makes this one insertion path rather than two: docking,
        // margins, out-of-flow positioning and the container's own layout are
        // all decided there, and a second copy of that reasoning would drift.
        if let Some(parent) = parent {
            let siblings = self.child_nodes(parent);
            let next = siblings
                .iter()
                .position(|c| *c == node)
                .and_then(|i| siblings.get(i + 1).copied());
            self.unlink_child(parent, node);
            if let Some(n) = self.nodes.get_mut(&node) {
                n.parent = None;
            }
            self.insert_before(parent, node, next);
        }

        // **`restyle_subtree` would do nothing here.** It applies what CHANGED,
        // and nothing about the cascade changed — the new control has simply
        // never been told any of it. Diffing against a blank set is what makes
        // "everything this element computes to" the change, which is the same
        // trick a first cascade plays on an element that has just been born.
        let computed = self.get_computed_style(node).clone();
        self.resolve_box_edges(node);
        for (property, value) in
            crate::css::changed_declarations(&crate::css::CssProperties::default(), &computed)
        {
            self.apply_style_property(node, property, &value);
        }

        // Everything the NODE says about the control has to be said again: the
        // widget that heard it the first time is gone. `value`, `checked`,
        // `disabled`, `placeholder`, `min`/`max` and the rest are all forwarded
        // to a control by `set_attribute` rather than kept on the node, so a
        // swap loses every one of them unless they are replayed. This is also
        // what makes the ORDER not matter: `value` then `type` and `type` then
        // `value` land the same way, which is the lesson `<progress>` taught
        // when `max` arriving second clamped a value that was already right.
        //
        // `type` is skipped because replaying it would re-enter here, and
        // `style` because the declaration store is not an attribute — it was
        // re-applied from the computed values just above, and re-parsing the
        // string would be a second path to the same place.
        let attributes: Vec<(String, String)> = self
            .node(node)
            .map(|n| {
                n.attributes
                    .iter()
                    .filter(|a| a.name != "type" && a.name != "style")
                    .map(|a| (a.name.clone(), a.value.clone()))
                    .collect()
            })
            .unwrap_or_default();
        for (name, value) in attributes {
            self.set_attribute(node, &name, &value);
        }
    }

    /// `parent.appendChild(child)`. Inserting is what makes an element render.
    /// A node has ONE parent, so appending it elsewhere moves it.
    ///
    /// Returns whether the child ended up in the parent. `false` is the
    /// spec's `HierarchyRequestError` case — a parent that cannot have
    /// children — surfaced rather than swallowed.
    pub fn append_child(&mut self, parent: NodeId, child: NodeId) -> bool {
        self.insert_before(parent, child, None)
    }

    /// `parent.insertBefore(child, reference)` — DOM §4.2.3.
    ///
    /// `reference` of `None` appends, which is what the spec says and what
    /// makes `appendChild` one line above rather than a second copy of all of
    /// this. A `reference` that is not a child of `parent` is the spec's
    /// `NotFoundError`: answered `false` rather than appended, because
    /// "somewhere in this parent" is not what the caller asked for.
    pub fn insert_before(
        &mut self,
        parent: NodeId,
        child: NodeId,
        before: Option<NodeId>,
    ) -> bool {
        // **A fragment inserts its CHILDREN and not itself** (DOM §4.2.1).
        // Inserting the fragment node would put a `#document-fragment` in the
        // tree — something no selector matches and no layout knows — and leave
        // the caller's nodes one level deeper than they asked for. Each child
        // goes before the SAME reference, which is what keeps them in order.
        //
        // Ahead of the checks below because they are about the node being
        // inserted, and after this line that node is a different one.
        if self.is_document_fragment(child) {
            let mut inserted = false;
            for node in self.child_nodes(child) {
                inserted |= self.insert_before(parent, node, before);
            }
            return inserted;
        }
        let parent = self.content_parent(parent, child);
        if let Some(reference) = before {
            if reference == child || !self.child_nodes(parent).contains(&reference) {
                return false;
            }
        }
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
            self.link_child(parent, child, before);
            self.sync_items(parent);
            return true;
        }

        // A comment, a CDATA section and a processing instruction have no
        // widget and no box, so they are subject to the one-parent rule and to
        // nothing else. Unlike a text node they are NOT gated on the parent
        // being able to draw a line — `<head><!--x--></head>` is as legal as
        // `<p><!--x--></p>` — and none of them triggers an inline rebuild,
        // because none contributes a run.
        if self.is_detached_data(child) {
            if let Some(previous) = self.nodes.get(&child).and_then(|n| n.parent) {
                self.unlink_child(previous, child);
            }
            if let Some(n) = self.nodes.get_mut(&child) {
                n.parent = Some(parent);
            }
            self.link_child(parent, child, before);
            return true;
        }

        // A text node has no widget to extract, so it is placed before the
        // check below rather than failing it. It is still subject to the
        // one-parent rule and still needs a container that can draw a line.
        if self.is_text_node(child) {
            // **A leaf control's text children are its caption.** A box draws
            // them as runs of its line; a `<button>` draws one string, and that
            // string is its content. Refusing them here made
            // `button.appendChild(createTextNode("7"))` and
            // `button.textContent = "7"` different — the spec defines the
            // second AS the first, so one of them rendering an empty button was
            // the tree and the widget disagreeing about the same fact.
            //
            // The document itself still refuses: root text has no box to belong
            // to, which `DocumentSink` already reports rather than drops.
            let box_ = self.holds_inline_content(parent);
            if !box_ && self.node(parent).is_none() {
                return false;
            }
            if let Some(previous) = self.nodes.get(&child).and_then(|n| n.parent) {
                self.unlink_child(previous, child);
                self.rebuild_inline_content(previous);
            }
            if let Some(n) = self.nodes.get_mut(&child) {
                n.parent = Some(parent);
            }
            self.link_child(parent, child, before);
            if box_ {
                self.rebuild_inline_content(parent);
                // **Text is what a column is measured FROM.** A `<td>` gets its
                // text as a TEXT NODE, and this branch returns before the one
                // that tells an enclosing table a cell changed — so a cell's
                // content never reached `column_widths` on the pass that used
                // it. The `else` arm above is safe already: `apply_text` ends
                // by asking the table.
                //
                // The tell was that a table needed TWO rows to size correctly.
                // Text landing after the table's last measurement leaves that
                // cell contributing nothing, so with two rows the other row
                // rescued the column and it looked correct; with one row both
                // columns fell back to an even split, the SHORTER text taking
                // the WIDER column because its min-content was the only
                // measurement that had arrived in time.
                self.relayout_enclosing_table(parent);
            } else {
                let text = self.text_content(parent);
                self.apply_text(parent, &text);
            }
            return true;
        }

        // **An inline ELEMENT inside a leaf control is its caption too.**
        //
        // The text-node branch above already says a leaf's character data is
        // what it draws. `<button><span>7</span></button>` is the same fact one
        // level up, and it was refused: a leaf cannot `add_child`, so the span
        // went to `detached` and the button rendered blank. Flutter spells every
        // label that way — `ElevatedButton(child: Text('7'))` — so an entire
        // keypad came out as empty rounded rectangles.
        //
        // Only for boxes that cannot hold children AND only for inline content:
        // a container keeps its real child, and a block-level child of a leaf is
        // a genuine structural error rather than a caption.
        if !self.holds_inline_content(parent)
            && self.node(parent).is_some()
            && self.is_inline_content(child)
        {
            if let Some(previous) = self.nodes.get(&child).and_then(|n| n.parent) {
                self.unlink_child(previous, child);
                self.rebuild_inline_content(previous);
            }
            if let Some(n) = self.nodes.get_mut(&child) {
                n.parent = Some(parent);
            }
            self.link_child(parent, child, before);
            // **A node that joined the tree has a computed style, whatever it
            // joined.** The run path below restyles here and this branch did
            // not, so anything nested under a leaf — `<strong><em>`, or the
            // `<span>` in `<button><span>7</span></button>` — was linked with
            // NO computed style at all: not inherited, not from the UA sheet.
            // `style_properties` is a stored read, so the answer was `None`
            // rather than a wrong value, and nothing downstream could tell the
            // difference between "not styled" and "not visited".
            self.restyle_subtree(child);
            self.restyle_positional_siblings(child);
            let text = self.text_content(parent);
            self.apply_text(parent, &text);
            return true;
        }

        let Some(widget) = self.extract_widget(child) else {
            return false;
        };

        // Unlink from the previous parent — the one-parent rule.
        if let Some(previous) = self.nodes.get(&child).and_then(|n| n.parent) {
            self.unlink_child(previous, child);
        }

        // **An inline element is not a box.** It joins the tree as a node — so
        // it keeps an id, a style and a `textContent` — but never as a child
        // widget: it contributes a styled RUN of its parent's text instead. Its
        // widget stays detached, which is also what keeps `set_text_content`
        // and the cascade working on it unchanged.
        if (self.is_text_node(child) || self.is_inline_content(child))
            && self.holds_inline_content(parent)
        {
            self.detached.insert(child, widget);
            if let Some(n) = self.nodes.get_mut(&child) {
                n.parent = Some(parent);
            }
            self.link_child(parent, child, before);
            self.restyle_subtree(child);
            self.restyle_positional_siblings(child);
            self.rebuild_inline_content(parent);
            // **Text is what a column is measured FROM**, so a table has to
            // look again for exactly the reason a new cell makes it look
            // again — and this branch returns before the one below that says
            // so. A `<td>` gets its text as a TEXT NODE, which is the only
            // kind of child that takes this path, so the effect was that no
            // cell's content ever reached `column_widths` on the pass that
            // used it.
            //
            // The tell was that a table needed TWO rows to size correctly. The
            // last row to be filled is the one whose text lands after the
            // table's final measurement, so with two rows the first one
            // rescued the column and the defect looked like a text-path
            // problem. With a single row nothing rescued it and both columns
            // fell back to an even split — the SHORTER text taking the wider
            // column, its min-content being the only measurement that had
            // arrived in time.
            self.relayout_enclosing_table(parent);
            return true;
        }

        // Where among the parent's WIDGETS this one goes. The DOM list is
        // already ordered by `link_child`; this is the same position expressed
        // for the container, and it is `None` for an append.
        let index = before.map(|reference| {
            self.child_nodes(parent)
                .iter()
                .position(|c| *c == reference)
                .unwrap_or(usize::MAX)
        });
        // **What kind of box is arriving, told BEFORE it arrives.**
        // `insert_widget` runs the container's layout as part of adding the
        // child, so a container told afterwards has already arranged it as a
        // block — and a block box is stretched to the content width. Nothing
        // put that width back: `layout_normal_flow` reads a child's CURRENT
        // size, so one pass under the wrong assumption is permanent, and a
        // `<button>` styled `width: 60px` came out 780 wide with its own
        // computed style still correctly reading `inline-block`.
        self.announce_child_display(parent, child);
        let inserted = self.insert_widget(parent, widget, index);
        match inserted {
            None => {
                self.link_child(parent, child, before);
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
                // And once more for margins, which the container also owns: a
                // `margin` set before the element joined its parent was told to
                // whichever container it did not yet have.
                if !self.box_edges(child).margin.is_zero() {
                    self.send_child_margin(child);
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
                // What an element inherits is decided by WHERE IT IS, so
                // joining a parent is a restyle. `Parent := Form` is the last
                // line of every VCL control's setup, long after the form
                // declared its font — without this the control was told
                // nothing and inheritance worked only for elements that
                // happened to be appended first.
                self.restyle_subtree(child);
                self.restyle_positional_siblings(child);
                // And what FORMATTING CONTEXT it lives in is decided the same
                // way — by where it is. The container has to be told what kind
                // of box just joined it before it can arrange it, and the box
                // has to be told what kind of container it is itself.
                self.apply_formatting(child);
                // …and joining a TABLE means the table has to look again: its
                // columns are measured from cell content, so a cell that
                // arrives after the table last laid out changes every column.
                self.relayout_enclosing_table(child);
                // Joining the body means joining its block flow, for the same
                // reason joining any other container re-runs that container's
                // layout. Without it a page appending elements with no
                // coordinates piled them all at the origin.
                if parent == DOCUMENT || Some(parent) == self.body {
                    self.relayout_body();
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
        index: Option<usize>,
    ) -> Option<Box<dyn PanelWidget>> {
        if parent == DOCUMENT {
            // The body positions absolutely — the endgame shape: an HTML page
            // whose controls carry `position:absolute` coordinates. Order is
            // not the form's to keep: `relayout_body` walks `child_nodes`,
            // which is the DOM's list, so an ordered insert at the root is
            // already honoured by the time this runs.
            let r = widget.rect();
            self.form.add_boxed_control(widget, r.x, r.y, r.w, r.h);
            return None;
        }
        let name = Self::widget_name(parent);
        let container = match self.detached_root(parent) {
            Some(root) => find_widget_mut(self.detached.get_mut(&root)?.as_mut(), &name),
            None => find_widget_mut(&mut self.form, &name),
        };
        match (container, index) {
            // A container that cannot honour an index hands the child back,
            // and `append_child`'s refusal path reports it. It must NOT quietly
            // append: the DOM would read back the requested order and the
            // window would show another.
            (Some(c), Some(index)) => c.insert_child(index, widget),
            (Some(c), None) => c.add_child(widget),
            (None, _) => Some(widget),
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

    /// `parent.replaceChild(new_child, old_child)` — DOM §4.2.3.
    ///
    /// Composed rather than open-coded: insert before the outgoing node, then
    /// remove it. Written out it would be a third copy of the branch ladder in
    /// `insert_before` — options, comments, text, inline, widget — and the
    /// copies would drift.
    ///
    /// Ordering matters and is the reason it is this way round: removing first
    /// would lose the POSITION, leaving nothing to insert before.
    pub fn replace_child(&mut self, parent: NodeId, new_child: NodeId, old_child: NodeId) -> bool {
        if new_child == old_child {
            return self.child_nodes(parent).contains(&old_child);
        }
        if !self.insert_before(parent, new_child, Some(old_child)) {
            return false;
        }
        self.remove_child(parent, old_child)
    }

    /// `node.cloneNode(deep)` — DOM §4.4.
    ///
    /// A copy of the node, **not in the document**: same tag, attributes and
    /// declarations, no parent, and — per spec — no children unless `deep`.
    ///
    /// Attributes and style are replayed through `set_attribute` /
    /// `set_style_property` rather than copied field-for-field, because those
    /// are where an attribute's MEANING lives: `value`, `checked`, `disabled`
    /// and `width` each reach a control or the box model, and a struct copy
    /// would duplicate the record while the clone's widget kept its defaults.
    pub fn clone_node(&mut self, node: NodeId, deep: bool) -> Option<NodeId> {
        let source = self.node(node)?.clone();
        let clone = match source.kind {
            NodeKind::Text => self.create_text_node(&source.data),
            NodeKind::Comment => self.create_comment(&source.data),
            NodeKind::CData => self.create_cdata_section(&source.data),
            NodeKind::ProcessingInstruction => {
                self.create_processing_instruction(&source.tag, &source.data)
            }
            // Cloning a fragment gives an EMPTY fragment; `deep` below fills
            // it. A fragment carries no name, no attributes and no data, so
            // there is nothing else of it to copy.
            NodeKind::DocumentFragment => self.create_document_fragment(),
            NodeKind::Element => {
                let clone = self.create_element_typed(&source.tag, &source.input_type);
                // The namespace travels with the clone — an element that lost
                // it would be in the wrong vocabulary the moment it was
                // inserted, and nothing about the tag would show it.
                if let Some(n) = self.nodes.get_mut(&clone) {
                    n.namespace = source.namespace.clone();
                }
                // Document order, which an attribute LIST has and a map did
                // not — two clones of one node are identical without a sort.
                let attributes = source.attributes.clone();
                for Attribute {
                    namespace,
                    name,
                    value,
                } in attributes
                {
                    if let Some(n) = self.nodes.get_mut(&clone) {
                        n.set_attribute(namespace, &name, &value);
                    }
                    self.set_attribute(clone, &name, &value);
                }
                let declarations: Vec<(String, String)> = source
                    .style
                    .iter()
                    .map(|(name, value)| (name.to_string(), value.to_string()))
                    .collect();
                for (property, value) in declarations {
                    self.set_style_property(clone, &property, &value);
                }
                clone
            }
        };
        if deep {
            for child in self.child_nodes(node) {
                if let Some(copy) = self.clone_node(child, true) {
                    self.append_child(clone, copy);
                }
            }
        }
        Some(clone)
    }

    /// `parent.removeChild(child)` — back out of the document, not destroyed.
    pub fn remove_child(&mut self, parent: NodeId, child: NodeId) -> bool {
        let parent = self.content_parent(parent, child);
        // **Being a child is the precondition — not owning a widget.**
        //
        // This used to bail when `extract_widget` came back empty, which reads
        // as a "not there" check and is not one: a text or comment node never
        // had a widget, because it is content of the box it sits in rather
        // than a box. So `removeChild(textNode)` unlinked nothing, returned
        // `false`, and said nothing — and `set_text_content`, which clears a
        // box by looping this over its children, left the old text behind and
        // appended the new beside it.
        if !self.child_nodes(parent).contains(&child) {
            return false;
        }
        let widget = self.extract_widget(child);
        self.unlink_child(parent, child);
        if let Some(n) = self.nodes.get_mut(&child) {
            n.parent = None;
        }
        // Kept for re-insertion, when there was one to keep.
        if let Some(widget) = widget {
            self.detached.insert(child, widget);
        }
        // Leaving the tree changes what this element inherits just as surely as
        // joining one does — it now inherits nothing. Without this it would
        // keep the colour and font of a parent it is no longer inside, and get
        // them back on re-insertion somewhere else.
        self.restyle_subtree(child);
        // And if it was a run of its parent's text, the parent's line is now a
        // different line.
        self.rebuild_inline_content(parent);
        true
    }

    /// `element.isConnected`.
    pub fn is_connected(&self, node: NodeId) -> bool {
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
        let name = self.fold_name(name);
        if node == DOCUMENT {
            if name == "title" {
                self.set_title(value);
            }
            self.nodes
                .entry(DOCUMENT)
                .or_default()
                .set_attribute(None, &name, value);
            return;
        }
        match name.as_str() {
            "value" => self.set_value(node, value),
            "checked" => self.set_checked(node, !value.eq_ignore_ascii_case("false")),
            // **`type` REALIZES the control**, it does not describe one.
            //
            // See [`Document::realize_control`] for why there was nothing here
            // and what that cost: every `<input>` a script built rendered as a
            // text box, because the type always arrived after the element did.
            "type" => {
                let now = self.fold_name(value.trim());
                if let Some((tag, was)) = self.node(node).map(|n| (n.tag.clone(), n.input_type.clone()))
                {
                    if was != now {
                        if let Some(n) = self.nodes.get_mut(&node) {
                            n.input_type = now.clone();
                        }
                        // Only a change of CONTROL is a rebuild. `text` to
                        // `search` is the same widget, and swapping it would
                        // throw away a caret and a selection for nothing.
                        if control_kind(&tag, &was) != control_kind(&tag, &now) {
                            self.realize_control(node);
                        }
                    }
                }
            }
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
            // **`width`/`height` are the BITMAP, not the box** — HTML §4.12.5.
            //
            // On a `<canvas>` these content attributes size the drawing
            // surface; CSS `width`/`height` size the box it is displayed in,
            // and the two are genuinely different numbers (a 640x480 bitmap
            // stretched into a 320x240 box is the ordinary way to draw at
            // double density). There was no arm at all, so `<canvas
            // width="640">` stored an inert attribute and the surface kept its
            // 300x150 default — which is what SDL, `CreateGraphics` and every
            // other painting caller sets first.
            //
            // The spec also says setting either attribute CLEARS the bitmap to
            // transparent black, even when the value does not change. That is
            // not a quirk to smooth over: callers rely on `canvas.width =
            // canvas.width` as the idiomatic full reset.
            //
            // On `<img>` and the other embedded elements the same attributes
            // are presentational hints that map to CSS instead, which is why
            // they are routed to the style store rather than to a bitmap.
            "width" | "height" => {
                let is_canvas = self
                    .node(node)
                    .map(|n| control_kind(&n.tag, &n.input_type) == "canvas")
                    .unwrap_or(false);
                if is_canvas {
                    if let Some(px) = self.px(node, &name, value) {
                        let mut rect = self.border_rect(node);
                        if name == "width" {
                            rect.w = px;
                        } else {
                            rect.h = px;
                        }
                        if let Some(w) = self.widget_mut(node) {
                            w.set_rect(rect);
                        }
                    }
                    if let Some(canvas) = self.canvas_mut(node) {
                        canvas.canvas_mut().clear();
                    }
                } else {
                    // A presentational hint is a declaration in the AUTHOR
                    // origin at zero specificity; recording it as an inline
                    // style is close enough here and keeps one code path,
                    // because nothing else can currently outrank it.
                    self.set_style_property(node, &name, &format!("{value}px"));
                }
            }
            // **`style=""` is a declaration BLOCK, not an opaque string.**
            //
            // CSSOM §6.1: the content attribute and `element.style` are the
            // same declarations seen two ways — writing the attribute writes
            // the store, and reading it back serialises the store. There was
            // no arm, so the attribute was recorded inert beside a store that
            // never heard about it: `<p style="color:red">` parsed, answered
            // `getAttribute("style")` correctly, and painted black.
            //
            // Returning early is the other half. Keeping a copy in
            // `attributes` would be the second storage this whole exercise
            // exists to remove, and the serialiser — which already writes the
            // store — would emit `style` twice.
            "style" => {
                for declaration in value.split(';') {
                    let Some((property, declared)) = declaration.split_once(':') else {
                        continue;
                    };
                    let property = property.trim().to_ascii_lowercase();
                    let declared = declared.trim();
                    if property.is_empty() || declared.is_empty() {
                        continue;
                    }
                    self.set_style_property(node, &property, declared);
                }
                return;
            }
            // **A span is a LAYOUT input, and layout is in the widget tree.**
            //
            // `colspan`/`rowspan` are content attributes, so they are already
            // stored on the node — but `flow_layout.rs` has no `Document` and
            // no `NodeId`, and the only channel from here to a widget is a
            // command. So the cell is told, the same way `min`/`max`/`src` are
            // told, and the table reads it back off the cell when it builds its
            // grid.
            //
            // A missing or unparseable span is 1, not 0: HTML §4.9.11 clamps to
            // at least one, and a zero would silently drop the cell out of every
            // column it should occupy.
            "colspan" | "rowspan" => {
                let span = value.trim().parse::<f64>().unwrap_or(1.0).max(1.0);
                self.command(
                    node,
                    &WidgetCommand::Custom(
                        format!("Set{}", capitalize(&name)),
                        CommandValue::Number(span),
                    ),
                );
                // A span changes the SHAPE of the grid, not just one cell —
                // every row after it shifts — so the table rebuilds it.
                self.relayout_enclosing_table(node);
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
            n.set_attribute(None, &name, value);
        }
        // **An attribute write is a RESTYLE.** `id`, `class` and every
        // `[attr]` selector match on attributes, so changing one changes what
        // the element — and everything under it, via descendant selectors —
        // computes to. Nothing re-ran the cascade here, so an element that
        // received its `class` AFTER being appended kept the style it had with
        // no class at all.
        //
        // Invisible until markup was PARSED: code that builds a tree by hand
        // tends to set attributes before appending, and the cascade then runs
        // once, correctly, at insertion. The parser does the opposite, so a
        // `<div class="side">` from `innerHTML` matched `div` and never
        // `.side` — a whole stylesheet silently not applying.
        // Only once the node is IN the tree. A detached element has no
        // ancestors, so a descendant selector cannot be judged yet — and
        // restyling early is actively harmful: the cascade would settle here,
        // and the restyle at insertion would then find NO difference and push
        // nothing to the container. Measured, by two grid tests: an item whose
        // `grid-column` was styled before it was appended lost its span.
        let connected = self
            .nodes
            .get(&node)
            .and_then(|n| n.parent)
            .is_some();
        if connected && !self.stylesheet.is_empty() {
            self.restyle_subtree(node);
        }
    }

    /// `element.getAttribute(qualifiedName)` — `None` when absent, which the
    /// surface turns into `null`.
    pub fn get_attribute(&mut self, node: NodeId, name: &str) -> Option<String> {
        let name = self.fold_name(name);
        // Live properties answer from the control, not from what was written.
        match name.as_str() {
            "value" => return Some(self.value(node)),
            "checked" => {
                return self.checked(node).then(|| "".to_string());
            }
            // The other half of the write above: the attribute IS the
            // serialised declaration store, so an element nobody styled has no
            // `style` attribute rather than an empty one.
            "style" => {
                return self
                    .style(node)
                    .filter(|s| !s.is_empty())
                    .map(|s| s.css_text());
            }
            _ => {}
        }
        self.nodes.get(&node)?.attribute(&name).map(str::to_string)
    }

    /// `element.getAttributeNames()` — DOM §4.9.
    ///
    /// **The names for which [`Document::get_attribute`] answers**, which is
    /// the only definition that cannot disagree with itself. It is not the
    /// stored list: an element's content attributes are SPLIT here. Everything
    /// a control owns — `value`, `checked`, `disabled` — is forwarded to the
    /// widget by [`Document::set_attribute`] and never reaches `attributes`,
    /// so enumerating that list alone would faithfully report `id`, `class` and
    /// `data-*` while silently omitting exactly the attributes that change.
    /// A reconciler built on it would look correct in testing and miss the
    /// live ones.
    ///
    /// The reconstructed names are therefore added back, and only when they
    /// actually answer — an unstyled element has no `style`, an empty field no
    /// `value` — so this does not put an attribute on every element in the
    /// document.
    ///
    /// ⚠ `disabled`/`hidden`/`min`/`max` are still missing, and not because
    /// they are skipped here: `get_attribute` has no arm for them either, so
    /// `setAttribute("disabled", "")` followed by `getAttribute("disabled")`
    /// already answers `null`. That is one bug in the read path, not two, and
    /// closing it there is what would make them appear here.
    pub fn get_attribute_names(&mut self, node: NodeId) -> Vec<String> {
        let mut names: Vec<String> = self
            .node(node)
            .map(|n| n.attributes.iter().map(|a| a.name.clone()).collect())
            .unwrap_or_default();
        for reconstructed in ["style", "value", "checked"] {
            if names.iter().any(|n| n == reconstructed) {
                continue;
            }
            // Empty is absent for these: a control reports `""` for a value
            // nobody set, and listing that would claim an attribute the element
            // does not have.
            if self
                .get_attribute(node, reconstructed)
                .is_some_and(|v| !v.is_empty() || reconstructed == "checked")
            {
                names.push(reconstructed.to_string());
            }
        }
        names
    }

    /// `element.setAttributeNS(namespace, qualifiedName, value)` — DOM §4.9.
    ///
    /// The qualified name is what serialises and what `getAttribute` finds;
    /// the namespace is what `getAttributeNS` matches on. Routed through
    /// `set_attribute` so a namespaced write still reaches the control that
    /// owns the attribute — a namespaced `value` is still a value.
    pub fn set_attribute_ns(&mut self, node: NodeId, namespace: &str, name: &str, value: &str) {
        self.set_attribute(node, name, value);
        let name = self.fold_name(name);
        if let Some(n) = self.nodes.get_mut(&node) {
            if let Some(attribute) = n.attributes.iter_mut().find(|a| a.name == name) {
                // Empty is `null`, per spec — not the empty string.
                attribute.namespace = (!namespace.is_empty()).then(|| namespace.to_string());
            }
        }
    }

    /// `element.getAttributeNS(namespace, localName)`.
    ///
    /// A different question from `getAttribute`, not a spelling of it:
    /// `xlink:href` and `href` share a local name and are two attributes, so
    /// this matches on the namespace and the LOCAL name while `getAttribute`
    /// matches the qualified one.
    pub fn get_attribute_ns(&self, node: NodeId, namespace: &str, local_name: &str) -> Option<String> {
        let namespace = (!namespace.is_empty()).then_some(namespace);
        self.node(node)?
            .attribute_ns(namespace, local_name)
            .map(str::to_string)
    }

    /// `element.removeAttribute(qualifiedName)`.
    pub fn remove_attribute(&mut self, node: NodeId, name: &str) {
        let name = self.fold_name(name);
        match name.as_str() {
            "disabled" => {
                self.command(node, &WidgetCommand::SetEnabled(true));
            }
            "hidden" => {
                self.command(node, &WidgetCommand::SetVisible(true));
            }
            "checked" => self.set_checked(node, false),
            // Removing `type` is not "no type": HTML §4.10.5.1.2 gives the
            // attribute a MISSING VALUE DEFAULT of `text`, so the element goes
            // back to being a text field and the control has to follow.
            "type" => {
                if let Some((tag, was)) = self.node(node).map(|n| (n.tag.clone(), n.input_type.clone()))
                {
                    if !was.is_empty() {
                        if let Some(n) = self.nodes.get_mut(&node) {
                            n.input_type.clear();
                        }
                        if control_kind(&tag, &was) != control_kind(&tag, "") {
                            self.realize_control(node);
                        }
                    }
                }
            }
            _ => {}
        }
        if let Some(n) = self.nodes.get_mut(&node) {
            n.remove_attribute(&name);
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
        // A custom property name is case-SENSITIVE and must not be folded; the
        // store draws that line, so the name goes in as written.
        let property = if crate::css::is_custom_property(property) {
            property.trim().to_string()
        } else {
            property.to_ascii_lowercase()
        };
        let before = self.get_computed_style(node).clone();
        // **A grid template is stored EXPANDED.** `repeat(7, 1fr)` and
        // `1fr 1fr 1fr 1fr 1fr 1fr 1fr` are the same template, and CSSOM
        // serialises the resolved track list — so normalising on the way in
        // means the store answers a track COUNT to anyone who splits it, which
        // is how a toolkit's `ColumnCount` reads back what it wrote. Leave an
        // unparseable value exactly as written: refusing to normalise is not
        // the same as refusing to record.
        let value = &match (property.as_str(), crate::css::parse_track_template(value)) {
            ("grid-template-columns" | "grid-template-rows", Some(template)) => {
                crate::css::track_template_css(&template)
            }
            _ => value.to_string(),
        }[..];
        self.record_style(node, &property, value);
        // The declaration that was just written is applied directly, because
        // not every property a widget acts on has a computed FIELD to be
        // diffed: `dock` is the toolkit's own and has none, and `border-*`,
        // `border-radius` and `flex-basis` are box values the diff does not
        // carry. Relying on the diff alone silently dropped them — a docked
        // panel simply never docked.
        //
        // `var()` is substituted on the way to the widget and NOT into the
        // store: `element.style` must read back the `var()` the author wrote
        // (CSSOM serialises the specified value), while the widget can only act
        // on a resolved one. An unresolvable reference applies as EMPTY rather
        // than returning, so the element takes the value it would have had
        // without the declaration instead of freezing on the last one.
        // The cascade first, because everything below reads it. `box_edges` is
        // resolved FROM the computed style, so resolving edges before the
        // cascade measured the element as it was before the write — every
        // margin, padding and percentage test failed on exactly that.
        //
        // No diff push for this element: the direct apply below is what tells
        // its widget. Pushing it a second time is not merely wasteful — `right`
        // recomputes an origin from the rect it just set, so applying it twice
        // walked the box by a padding each time.
        self.resolve_computed(node);
        // `display` and `position` decide whether this element is a box at all,
        // so the answer is re-checked before anything is applied to a widget
        // that may not be the right kind of thing yet.
        self.reconcile_box_or_run(node);
        self.resolve_box_edges(node);
        let resolved = self.resolve_value(node, value).unwrap_or_default();
        self.apply_style_property(node, &property, &resolved);
        // The cascade is then the DESCENDANTS' business: what the write
        // propagates, which is where the guessing used to be.
        //
        // **Three special cases used to live here and are now consequences.**
        // A custom property needed its own branch because it is not a property
        // any widget acts on — it is a VALUE other declarations read, so the
        // write does nothing here and everything to whoever referenced it, on
        // this element and every descendant. An inherited property needed
        // another, because writing `color` on a form is the whole of how a VCL
        // `Font`/`Color` reaches the controls inside it. And `var()` needed
        // substituting on the way to the widget but NOT into the store, since
        // `element.style` must read back the `var()` the author wrote while the
        // widget can only act on a resolved one.
        //
        // All three are now what re-resolving the subtree and diffing already
        // does. Nothing asks what KIND of property was written, which is the
        // point: a cascade resolves, and what a declaration reaches is a
        // property of the cascade rather than a list to keep in step.
        //
        // On this element, everything the write moved EXCEPT the write itself.
        // A `--pad: ""` is the case that needs it: the declaration written is a
        // custom property, which no widget acts on, while the `padding:
        // var(--pad)` it invalidated sits on this same element and has to be
        // taken back.
        let own: Vec<(&'static str, String)> =
            crate::css::changed_declarations(&before, self.get_computed_style(node))
                .into_iter()
                .filter(|(name, _)| *name != property.as_str())
                .collect();
        let moved = !own.is_empty() || self.get_computed_style(node) != &before;
        for (name, value) in own {
            self.apply_style_property(node, name, &value);
        }
        // This element's own line, if it has one. Its text nodes declare
        // nothing, so their runs are built from THIS element's computed style —
        // and the write above went to the widget, not to the runs. The
        // descendant walk below cannot see it either, because the box is not
        // one of its own descendants.
        if moved && self.has_inline_content(node) {
            self.rebuild_inline_content(node);
        }
        for child in self.child_nodes(node) {
            self.restyle_subtree(child);
            self.restyle_positional_siblings(child);
        }
        // **A cell's style is a question for the TABLE.** Padding, borders and
        // font all feed the cell's min- and max-content width, and a column's
        // width is computed from every cell in it — so restyling one cell can
        // move every column and every row. Nothing else notices: the widget
        // took the declaration, but the box that decides its width is several
        // levels up and was never asked to look again.
        if moved {
            self.relayout_enclosing_table(node);
        }
    }

    /// Substitute any `var()` in a declared value, resolving names against
    /// `node` and its ancestors.
    fn resolve_value(&self, node: NodeId, value: &str) -> Option<String> {
        if !crate::css::references_var(value) {
            return Some(value.to_string());
        }
        crate::css::substitute_vars(value, &|name| self.custom_property(node, name))
    }

    /// What `--name` holds for this element.
    ///
    /// Custom properties INHERIT, so a name not declared here is asked of the
    /// parent, and so on to the root. That is the whole reason a theme can be
    /// declared once on the document and read by every control under it.
    fn custom_property(&self, node: NodeId, name: &str) -> Option<String> {
        let mut current = Some(node);
        while let Some(id) = current {
            if let Some(style) = self.style(id) {
                let declared = style.get(name);
                if !declared.is_empty() {
                    return Some(declared.to_string());
                }
            }
            current = self.nodes.get(&id).and_then(|n| n.parent);
        }
        None
    }


    /// Translate one declaration into whatever the widget can act on.
    ///
    /// Split from [`Document::set_style_property`] because the UA stylesheet
    /// applies declarations WITHOUT recording them in `element.style` — a UA
    /// rule is not an inline style, and serialising one into `style=""` would
    /// put `display:none` on every `<script>` in the output. The cascade is
    /// UA < author, and it falls out of the UA layer being applied first, at
    /// creation, before any program can write.
    fn apply_style_property(&mut self, node: NodeId, property: &str, value: &str) {
        let px = self.px(node, property, value);
        match property {
            "left" | "top" | "width" | "height" => {
                // A width arriving LATER changes whether the container may
                // stretch this box, and only the container can act on that.
                if matches!(property, "width" | "height") && node != DOCUMENT {
                    let horizontal = property == "width";
                    let mode = self.axis_mode(node, horizontal);
                    let verb = if horizontal {
                        "SetChildWidthMode"
                    } else {
                        "SetChildHeightMode"
                    };
                    self.tell_container(node, verb, mode);
                    // **One fact, both ends.** The container needs it to decide
                    // whether to stretch this box; the box needs it to decide
                    // whether its own main size is definite before flexing its
                    // items. Same answer, sent to both, so they cannot disagree.
                    if !horizontal {
                        self.command(
                            node,
                            &WidgetCommand::Custom(
                                "SetHeightMode".into(),
                                CommandValue::Text(mode.to_string()),
                            ),
                        );
                    }
                }
                // **A percentage with nothing to be a percentage OF writes no
                // pixels.** The declaration is already recorded, and
                // `size_subtree` re-resolves it once the box has a containing
                // block; writing a number now would put the viewport's size on
                // a box that is still detached, and — because a declared height
                // is not overruled by content — that wrong number would then
                // survive every later pass. A Flutter button came out the full
                // height of the window this way, in a `Padding` that never had
                // a height to give it.
                if value.trim().ends_with('%')
                    && matches!(property, "width" | "height")
                    && node != DOCUMENT
                    && !self.percentage_resolves(node, property == "width")
                {
                    return;
                }
                // `left`/`width` measure across, `top`/`height` down — which is
                // the axis a percentage refers to.
                let horizontal = matches!(property, "left" | "width");
                let Some(px) = self.resolve_length(node, value, horizontal) else {
                    return;
                };
                if node == DOCUMENT {
                    let r = self.form.rect();
                    let r = apply_axis(r, &property, px);
                    <Form as PanelWidget>::set_rect(&mut self.form, r);
                    return;
                }
                // `box-sizing` decides what the number MEASURES. The widget's
                // rect is the border box, so a `content-box` width has to grow
                // by the edges that sit outside the content — which is the
                // whole property, and the reason it cannot be answered without
                // a box model.
                let px = match property {
                    "width" | "height" if self.uses_content_box(node) => {
                        let edges = self.box_edges(node);
                        let extra = if property == "width" {
                            edges.border.horizontal() + edges.padding.horizontal()
                        } else {
                            edges.border.vertical() + edges.padding.vertical()
                        };
                        px + extra
                    }
                    _ => px,
                };
                let clamped = self.clamp_to_constraints(node, &property, px);
                // `left`/`top` are measured from the containing block, not from
                // the document. A rect is in form coordinates, so the block's
                // origin has to be added back — otherwise a child at (20, 50)
                // inside a panel at (10, 10) lands at (20, 50) on the BODY.
                // Invisible whenever the container sits at the origin, which is
                // why the calculator never showed it.
                //
                // CSS 2.1 §10.3.7/§10.6.4: for an absolutely positioned box the
                // offsets place its MARGIN edge, so the border box — which is
                // what a rect is — sits one margin further in. In flow the
                // margin is the container's business instead, and adding it
                // here would double-count what `SetChildMargin` already said.
                // Without this, `SetChildMargin` reaching a form — whose every
                // child is positioned and which therefore has no flow to apply
                // it to — meant a body-level margin was simply dropped.
                //
                // The question is `positions_itself`, NOT whether `position:
                // absolute` was declared. Only containers declare it; an
                // ordinary VCL control arrives with `left`/`top` and nothing
                // else and is inferred out of flow, so asking for the
                // declaration would have made this branch dead for every
                // control in the corpus while both unit tests still passed.
                let margin = if self.positions_itself(node) {
                    self.box_edges(node).margin
                } else {
                    crate::css::Edges::default()
                };
                let origin = match property {
                    "left" => self.containing_block(node).x + margin.left,
                    "top" => self.containing_block(node).y + margin.top,
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
                if matches!(property, "left" | "top") {
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
                    // **`block` did NOT leave it as it was.** This arm used to
                    // say so and stop, on the reasoning that a flex container
                    // arranges children along an axis and so does the flow
                    // panel. But "as it was" was flex with `align-items:
                    // stretch`, so every block box in the engine — every
                    // `<div>`, every `<p>`, the body itself — laid its children
                    // out as a stretched flex column. A row of `<button>`s came
                    // out as a column of full-width bars, and the buttons'
                    // own correctly-computed `inline-block` had no reader.
                    self.apply_formatting(node);
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
            // **`gap` is two properties, and the longhands are not the same
            // one.** All three used to send `SetGap`, which set a single
            // `spacing` scalar — so `row-gap: 20px; column-gap: 4px` gave
            // whichever declaration arrived last on BOTH axes.
            // `justify-items` is the container's answer for the inline axis;
            // `justify-self` is one child overruling it. Same split as
            // `align-items`/`align-self`, on the other axis — grid has two and
            // flex only ever needed one, which is why these had no route.
            "justify-items" => {
                self.command(
                    node,
                    &WidgetCommand::Custom(
                        "SetJustifyItems".into(),
                        CommandValue::Text(value.trim().to_ascii_lowercase()),
                    ),
                );
            }
            "justify-self" => {
                self.tell_container(node, "SetChildJustifySelf", &value.trim().to_ascii_lowercase());
            }
            "align-content" => {
                self.command(
                    node,
                    &WidgetCommand::Custom(
                        "SetAlignContent".into(),
                        CommandValue::Text(value.trim().to_ascii_lowercase()),
                    ),
                );
            }
            "gap" | "row-gap" | "column-gap" => {
                let Some(px) = px else { return };
                let verb = match property {
                    "row-gap" => "SetRowGap",
                    "column-gap" => "SetColumnGap",
                    _ => "SetGap",
                };
                self.command(
                    node,
                    &WidgetCommand::Custom(verb.into(), CommandValue::Number(px as f64)),
                );
            }
            // The table's own three. `border-collapse` and `border-spacing`
            // INHERIT (§17.6), so they reach every cell too — harmlessly, since
            // only a box running the table context reads them, and it is what
            // lets a cell find the border model without walking up to its table.
            "border-spacing" => {
                let Some(px) = px else { return };
                self.command(
                    node,
                    &WidgetCommand::Custom(
                        "SetBorderSpacing".into(),
                        CommandValue::Number(px as f64),
                    ),
                );
            }
            "border-collapse" | "table-layout" => {
                let verb = if property == "border-collapse" {
                    "SetBorderCollapse"
                } else {
                    "SetTableLayout"
                };
                self.command(
                    node,
                    &WidgetCommand::Custom(verb.into(), CommandValue::Text(value.to_string())),
                );
            }
            // `z-index` is a PAINT decision, so it goes to whoever paints the
            // children — and it deliberately does nothing for a static box,
            // which is CSS's rule and falls out of the container sorting only
            // its positioned children.
            // **The border, sent whole on any change to any side.** A border is
            // three axes over four sides and CSS lets a rule touch any one of
            // them; rebuilding the whole box from the computed style means the
            // widget never has to hold a half-updated border, and a shorthand
            // and a longhand arrive identically.
            //
            // Nothing sent these before. `box_edges` read the widths for
            // LAYOUT, so a declared border moved the content box and painted
            // nothing at all.
            "border-top-width" | "border-right-width" | "border-bottom-width"
            | "border-left-width" | "border-top-style" | "border-right-style"
            | "border-bottom-style" | "border-left-style" | "border-top-color"
            | "border-right-color" | "border-bottom-color" | "border-left-color" => {
                self.send_border(node);
                // Widths are part of the box, so the cached edges and the
                // element's own geometry both move with them.
                self.resolve_box_edges(node);
            }
            // Where this item sits in its parent's grid — a per-child fact, so
            // it goes to the container that places it.
            "grid-column" | "grid-row" | "grid-column-start" | "grid-column-end"
            | "grid-row-start" | "grid-row-end" => {
                self.send_grid_placement(node);
            }
            // The grid template. Sent whole, re-parsed by the widget, so the
            // track grammar has exactly one implementation.
            "grid-template-columns" => {
                self.command(
                    node,
                    &WidgetCommand::Custom(
                        "SetGridColumns".into(),
                        CommandValue::Text(value.into()),
                    ),
                );
            }
            "grid-template-rows" => {
                self.command(
                    node,
                    &WidgetCommand::Custom("SetGridRows".into(), CommandValue::Text(value.into())),
                );
            }
            "grid-auto-flow" => {
                self.command(
                    node,
                    &WidgetCommand::Custom(
                        "SetGridAutoFlow".into(),
                        CommandValue::Text(value.trim().to_ascii_lowercase()),
                    ),
                );
            }
            "grid-auto-rows" | "grid-auto-columns" => {
                let verb = if property == "grid-auto-rows" {
                    "SetGridAutoRows"
                } else {
                    "SetGridAutoColumns"
                };
                self.command(
                    node,
                    &WidgetCommand::Custom(verb.into(), CommandValue::Text(value.into())),
                );
            }
            "grid-template-areas" => {
                self.command(
                    node,
                    &WidgetCommand::Custom("SetGridAreas".into(), CommandValue::Text(value.into())),
                );
            }
            // The NAME form of `grid-area`. The four-line form writes the
            // longhands and arrives through the placement arm above.
            "grid-area" if !value.contains('/') => {
                self.tell_container(node, "SetChildGridName", value.trim());
            }
            // **Flex item sizing, which reached no widget at all.** Both parse
            // into `css.rs` and had no arm here, so `flex-basis` and
            // `flex-shrink` were stored and never asked for. Addressed to the
            // container for the same reason `flex-grow` is: it is the container
            // that distributes space.
            "flex-basis" => {
                let Some(px) = px else { return };
                self.tell_container(node, "SetChildFlexBasis", &px.to_string());
            }
            "flex-shrink" => {
                let Some(shrink) = value.trim().parse::<f32>().ok() else {
                    return;
                };
                self.tell_container(node, "SetChildFlexShrink", &shrink.to_string());
            }
            "z-index" => {
                self.tell_parent(node, "SetChildZ", value.trim().to_string());
            }
            // Margin is the child's claim on space between it and its siblings,
            // so it goes to the CONTAINER. Unlike padding, nothing about the
            // element's own box changes — which is why this sends and returns
            // rather than also touching the widget.
            "margin"
            | "margin-top"
            | "margin-right"
            | "margin-bottom"
            | "margin-left"
            | "margin-block-start"
            | "margin-inline-end"
            | "margin-block-end"
            | "margin-inline-start" => {
                self.send_child_margin(node);
                // A positioned box's offsets place its margin edge, so a margin
                // arriving AFTER `left` moves it — the same ordering problem
                // `reapply_constrained_axis` solves for `min-width`. A box in
                // flow needs nothing: the container was just told, and the
                // container is what acts.
                if self.positions_itself(node) {
                    self.reapply_positioned_offsets(node);
                }
                self.relayout_parent(node);
            }
            // Every spelling of padding — the shorthand, the four longhands and
            // their logical aliases — reaches the container as the same four
            // resolved numbers, because `box_edges` has already read the whole
            // declaration store. This arm used to flatten the shorthand to its
            // LARGEST edge, which made `padding: 0 40px` a 40px inset on all
            // four sides; a `<ul>`'s marker indent cannot be spelled that way.
            "padding"
            | "padding-top"
            | "padding-right"
            | "padding-bottom"
            | "padding-left"
            | "padding-block-start"
            | "padding-inline-end"
            | "padding-block-end"
            | "padding-inline-start" => {
                let p = self.box_edges(node).padding;
                self.command(
                    node,
                    &WidgetCommand::Custom(
                        "SetPadding".into(),
                        CommandValue::Text(format!("{},{},{},{}", p.top, p.right, p.bottom, p.left)),
                    ),
                );
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
            // The rest of the font. These reached nothing at all before: every
            // draw site named its own family and there was no channel for
            // weight or slant, so `b, strong { font-weight: bold }` was a
            // declaration with nowhere to arrive.
            //
            // `font-family` is a LIST in CSS and the first name wins here —
            // the fallback chain needs a font database to walk, and picking the
            // head is what `css.rs` already parses.
            "font-family" => {
                let family = value.split(',').next().unwrap_or("").trim().trim_matches(
                    |c| c == '"' || c == '\'',
                );
                self.command(
                    node,
                    &WidgetCommand::Custom(
                        "SetFontFamily".into(),
                        CommandValue::Text(family.to_string()),
                    ),
                );
            }
            "font-weight" => {
                self.command(
                    node,
                    &WidgetCommand::Custom(
                        "SetFontWeight".into(),
                        CommandValue::Text(value.trim().to_string()),
                    ),
                );
            }
            "font-style" => {
                self.command(
                    node,
                    &WidgetCommand::Custom(
                        "SetFontStyle".into(),
                        CommandValue::Text(value.trim().to_string()),
                    ),
                );
            }
            // Where the text sits in the control's own box. Inherited, so a
            // form declaring it once reaches every label under it — which is
            // also why the value handed over is the CSS keyword rather than a
            // parsed enum: `TextAlign::from_css` is the one place a keyword
            // becomes an alignment.
            "text-align" => {
                self.command(
                    node,
                    &WidgetCommand::Custom(
                        "SetTextAlign".into(),
                        CommandValue::Text(value.trim().to_string()),
                    ),
                );
            }
            // Whether the box's text may BREAK. The shaper already takes an
            // optional wrap width and `nowrap` is that width being absent, so
            // the box only has to be told which it is.
            "white-space" => {
                self.command(
                    node,
                    &WidgetCommand::Custom(
                        "SetWhiteSpace".into(),
                        CommandValue::Text(value.trim().to_ascii_lowercase()),
                    ),
                );
            }
            "text-decoration" | "text-decoration-line" => {
                self.command(
                    node,
                    &WidgetCommand::Custom(
                        "SetTextDecoration".into(),
                        CommandValue::Text(value.trim().to_string()),
                    ),
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
        let (left, top) = self
            .style(node)
            .map(|s| (s.get("left").to_string(), s.get("top").to_string()))
            .unwrap_or_default();
        let (dx, dy) = (self.px(node, "left", &left), self.px(node, "top", &top));
        let placement = format!(
            "relative:{},{}",
            dx.unwrap_or(0.0),
            dy.unwrap_or(0.0)
        );
        self.send_child_placement(node, placement);
    }

    fn send_child_placement(&mut self, node: NodeId, placement: String) {
        self.tell_parent(node, "SetChildFlow", placement);
    }

    /// The child's margins, told to whoever arranges it.
    ///
    /// A margin is space BETWEEN siblings, so only the container can honour it
    /// — the same reason `out_of_flow`, `relative_offset` and `order` all live
    /// with the container rather than the child. A child cannot reserve space
    /// it does not occupy.
    fn send_child_margin(&mut self, node: NodeId) {
        let m = self.box_edges(node).margin;
        self.tell_parent(
            node,
            "SetChildMargin",
            format!("{},{},{},{}", m.top, m.right, m.bottom, m.left),
        );
    }

    /// Address a per-child fact to the parent, which is the widget that acts on
    /// it. The document's own children are the form's, which is not a node.
    fn tell_parent(&mut self, node: NodeId, command: &str, value: String) {
        let Some(parent) = self.nodes.get(&node).and_then(|n| n.parent) else {
            return;
        };
        let spec = format!("{}={}", Self::widget_name(node), value);
        let cmd = WidgetCommand::Custom(command.into(), CommandValue::Text(spec));
        if parent == DOCUMENT {
            self.form.handle_command(&cmd);
            return;
        }
        self.command(parent, &cmd);
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
    ///
    /// The rectangle is the ancestor's **padding box**, not its border box —
    /// CSS 2.1 §10.1. A child at `left: 0` starts inside the border and outside
    /// nothing else, so the two answers differ by exactly the border width and
    /// coincide wherever it is zero.
    fn containing_block(&mut self, node: NodeId) -> LayoutRect {
        if self.declared_position(node) == "fixed" {
            return self.form.rect();
        }
        // **Only an absolutely positioned box looks for a positioned ancestor**
        // — CSS 2.1 §10.1. For a `static` or `relative` box the containing
        // block is simply the nearest block container ancestor's CONTENT edge,
        // which is its parent, and the walk below would sail straight past it.
        //
        // Applying the absolute rule to everything is what made `height: 100%`
        // on a Flutter button resolve against the VIEWPORT: nothing in the tree
        // was positioned, so every percentage — on a button four levels down —
        // answered 800x600 and the button covered the window.
        if self.declared_position(node) != "absolute" {
            let parent = self.nodes.get(&node).and_then(|n| n.parent);
            return match parent {
                Some(p) if p != DOCUMENT && self.node(p).is_some() => self.content_rect(p),
                // A child of the document resolves against the viewport, which
                // is the initial containing block — the same answer a browser
                // gives for a child of `<body>` with no sized ancestor.
                _ => self.form.rect(),
            };
        }
        let mut cursor = self.nodes.get(&node).and_then(|n| n.parent);
        while let Some(parent) = cursor {
            if parent == DOCUMENT {
                break;
            }
            if !matches!(self.declared_position(parent).as_str(), "" | "static") {
                if self.widget_mut(parent).is_some() {
                    return self.padding_rect(parent);
                }
            }
            cursor = self.nodes.get(&parent).and_then(|n| n.parent);
        }
        // No positioned ancestor: the initial containing block, which is the
        // viewport. A browser answers the same.
        self.form.rect()
    }

    fn declared_position(&self, node: NodeId) -> String {
        // Cascade first, for the same reason as `positions_itself`: a rule in a
        // stylesheet is as much an answer to "how is this box positioned" as an
        // inline declaration, and only one of the two used to be heard.
        if let Some(position) = self.get_computed_style(node).position {
            return position.as_css().to_string();
        }
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
    /// What `em` and `rem` mean for this element.
    ///
    /// `em` is the element's own computed `font-size`, **except on `font-size`
    /// itself**, where the value cannot be its own basis and the parent's is
    /// used — the same exception the cascade makes, stated the same way, so the
    /// value pushed to a widget cannot disagree with the value in the store.
    fn font_context(&self, node: NodeId, property: &str) -> crate::css::FontContext {
        let own = self.get_computed_style(node).font_size;
        let parent = self
            .node(node)
            .and_then(|n| n.parent)
            .and_then(|p| self.get_computed_style(p).font_size);
        let em = if property == "font-size" { parent } else { own.or(parent) };
        crate::css::FontContext {
            em: em.unwrap_or(16.0),
            rem: self.document_computed.font_size.unwrap_or(16.0),
        }
    }

    /// Pixels for a declaration on `node`, with font-relative units resolved.
    ///
    /// **The second copy of the length parser, closed.** This path hardcoded
    /// `em` to 16px and treated `rem` as a synonym for it — so the widget could
    /// be pushed a font-relative value the computed store had resolved
    /// differently, which is the same "two parsers, one question" shape the
    /// colour parser had.
    fn px(&self, node: NodeId, property: &str, value: &str) -> Option<f32> {
        crate::css::parse_length_in(value, self.font_context(node, property))
            .and_then(crate::css::Length::px)
    }

    fn resolve_length(&mut self, node: NodeId, value: &str, horizontal: bool) -> Option<f32> {
        if let Some(px) = self.px(node, "", value) {
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

    /// Re-apply a declared `left`/`top` — used when a margin arrives after the
    /// offset it displaces. Sibling of [`Document::reapply_constrained_axis`].
    fn reapply_positioned_offsets(&mut self, node: NodeId) {
        for property in ["left", "top"] {
            let declared = self
                .style(node)
                .map(|s| s.get(property).to_string())
                .filter(|d| !d.is_empty());
            if let Some(declared) = declared {
                self.apply_style_property(node, property, &declared);
            }
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
        let spec = format!("{}={}", Self::widget_name(node), value);
        self.command_container(
            parent,
            &WidgetCommand::Custom(verb.to_string(), CommandValue::Text(spec)),
        );
    }

    /// Send a container-directed command, INCLUDING when the container is the
    /// document itself.
    ///
    /// `command` addresses a widget by name through `form.send_command`, and the
    /// root is not in its own child list — so a command aimed at `DOCUMENT`
    /// found nothing and did nothing. Both callers here used to bail on
    /// `parent == DOCUMENT` rather than route, which read as a decision and was
    /// a hole: `document.body` IS the root node, so EVERY per-child layout fact
    /// about an element appended to the body — `display`, the flex trio,
    /// `align-self`, `justify-self`, `order`, grid placement, the axis modes —
    /// was dropped in silence. A page built the web way, by appending to
    /// `document.body`, could only ever stack its children down the left edge,
    /// because the root was never told any of them was inline-level.
    ///
    /// The routing already existed twice in this file (`set_child_*` sends to
    /// `form.handle_command`, `relayout_parent` calls `relayout_body`); these
    /// two just predate it.
    fn command_container(&mut self, parent: NodeId, cmd: &WidgetCommand) {
        if parent == DOCUMENT {
            self.form.handle_command(cmd);
            return;
        }
        self.command(parent, cmd);
    }

    /// Record a child's grow weight with its container.
    fn set_child_flex(&mut self, node: NodeId, grow: f32) {
        self.tell_container(node, "SetChildFlex", &grow.to_string());
    }

    /// **Tell the widgets what `display` means** — both halves of it.
    ///
    /// `display` says two different things depending on who is asking, and the
    /// engine was only ever told one of them:
    ///
    /// - To the box ITSELF it names the formatting context it establishes for
    ///   its children. `flex` is the flex algorithm; everything else is normal
    ///   flow. Until this existed every container ran flex, so a `<div>` — a
    ///   block box — arranged its children as a stretched flex column.
    /// - To the box's PARENT it says whether this box takes a row of its own or
    ///   shares a line with its siblings. Only the parent can act on that,
    ///   which is why it is sent there.
    ///
    /// `none` is neither, and it is not a `Display` here at all — visibility is
    /// its own flag, set where the property is written. This asks only what
    /// kind of box a displayed box is.
    /// Tell `parent` what kind of box `child` is, naming the parent rather than
    /// reading it off the child.
    ///
    /// The child is not linked yet at the moment this matters — it is about to
    /// be handed to the container, and the container arranges it on the way in.
    /// [`Document::tell_container`] asks the child who its parent is, which is
    /// still nobody.
    fn announce_child_display(&mut self, parent: NodeId, child: NodeId) {
        let Some(display) = self.get_computed_style(child).display else {
            return;
        };
        let spec = format!("{}={}", Self::widget_name(child), display.as_css());
        self.command_container(
            parent,
            &WidgetCommand::Custom("SetChildDisplay".into(), CommandValue::Text(spec)),
        );
        // **`flex-shrink`'s initial value is 1 — for CSS.** The panel defaults
        // to 0 because Flutter's `Row` overflows rather than shrinking, and
        // Flutter reaches the widget by KIND with no element and no cascade.
        // An ELEMENT is a CSS box, so it gets CSS's initial value, and the two
        // defaults stop fighting by never meeting.
        //
        // Named parent, not `tell_container`: nothing is linked yet at this
        // point, so asking the child who its parent is answers nobody and the
        // send would silently do nothing.
        let shrink = self.get_computed_style(child).flex_shrink.unwrap_or(1.0);
        for (verb, mode) in [
            ("SetChildFlexShrink", shrink.to_string()),
            ("SetChildWidthMode", self.axis_mode(child, true).to_string()),
            ("SetChildHeightMode", self.axis_mode(child, false).to_string()),
        ] {
            let spec = format!("{}={}", Self::widget_name(child), mode);
            self.command(
                parent,
                &WidgetCommand::Custom(verb.into(), CommandValue::Text(spec)),
            );
        }
    }

    /// Whether a box's width is `auto` — the one fact that decides if normal
    /// flow may stretch it.
    ///
    /// CSS fills the content width only when the width is `auto`; a declared
    /// width is kept. The container knew each child's `display` and not this,
    /// so a block-level child with `width: 200px` was stretched to its
    /// container anyway. `fill_available_width` makes exactly this distinction
    /// for the body's children; this is the same question one level down.
    fn axis_mode(&self, node: NodeId, horizontal: bool) -> &'static str {
        let props = self.get_computed_style(node);
        let declared = if horizontal {
            props.width.is_some()
        } else {
            props.height.is_some()
        };
        if declared {
            // …unless it is a PERCENTAGE of something indefinite, which is
            // `auto` — §10.5. The declaration exists but resolves to nothing,
            // so a box under it must not resolve against it either.
            if !horizontal && !self.percentage_height_resolves(node) {
                return "auto";
            }
            return "declared";
        }
        // **A flex item's main size is DEFINITE after flexing** — §9.9. A box
        // that grew into a share of a container that had a height to share is
        // as definite as one that declared a length, and its own children may
        // resolve against it: that is how `Scaffold { flex: 1 }` passes the
        // window's height down to the `Expanded` rows without every level
        // repeating `height`.
        //
        // Without it the chain broke at the first box whose size came from
        // flexing rather than from a declaration — the container said `auto`,
        // the flow sized its children to content, and the space it had just
        // been given went nowhere.
        if !horizontal && self.flexed_block_size_is_definite(node) {
            return "declared";
        }
        "auto"
    }

    /// Whether a declared `height` actually resolves to a length.
    ///
    /// True for any non-percentage declaration, and for a percentage only when
    /// the containing block's own block size is definite. `height: 100%` of a
    /// parent that is still sizing to its content is `auto`, and a box in that
    /// state must be measured from its content like any other — otherwise it
    /// keeps whatever pixel value the percentage happened to resolve to when it
    /// was first written, which for a detached element is the viewport's.
    ///
    /// That is what made a Flutter `Padding` in a plain `Column` — a LOOSE
    /// constraint, so its child sizes itself — hand the tictactoe "New Game"
    /// button the full height of the window.
    fn percentage_height_resolves(&self, node: NodeId) -> bool {
        let is_percent = self
            .style(node)
            .map(|s| s.get("height").trim().ends_with('%'))
            .unwrap_or(false);
        if !is_percent {
            return true;
        }
        self.percentage_resolves(node, false)
    }

    /// Whether a percentage on one axis has a containing block to resolve
    /// against **right now**. A detached box has none at all; on the block axis
    /// an attached one still needs its parent's own height to be definite.
    fn percentage_resolves(&self, node: NodeId, horizontal: bool) -> bool {
        match self.nodes.get(&node).and_then(|n| n.parent) {
            // The viewport is the initial containing block, definite on both
            // axes.
            Some(p) if p == DOCUMENT => true,
            // A block container's width always resolves; its height only if it
            // has one.
            Some(p) => horizontal || self.axis_mode(p, false) == "declared",
            None => false,
        }
    }

    /// Whether this box lays its children out in tracks it decides itself —
    /// a flex or grid formatting context, where the item's size comes from the
    /// container's algorithm rather than from the item filling its parent.
    fn establishes_independent_track_layout(&self, node: NodeId) -> bool {
        use crate::css::Display;
        matches!(
            self.get_computed_style(node).display,
            Some(Display::Flex)
                | Some(Display::InlineFlex)
                | Some(Display::Grid)
                | Some(Display::InlineGrid)
        )
    }

    /// Whether this box's block size came from **flexing** a container that had
    /// a definite one to divide.
    ///
    /// Three things have to hold, and the third is recursive: the parent
    /// establishes a flex context, its main axis is the block axis (a column —
    /// in a row the block axis is the CROSS one, where stretching is not the
    /// same guarantee), and the parent's own block size is itself definite.
    fn flexed_block_size_is_definite(&self, node: NodeId) -> bool {
        let Some(parent) = self.nodes.get(&node).and_then(|n| n.parent) else {
            return false;
        };
        if parent == DOCUMENT {
            return false;
        }
        let container = self.get_computed_style(parent);
        let grid = matches!(
            container.display,
            Some(crate::css::Display::Grid) | Some(crate::css::Display::InlineGrid)
        );
        if !grid
            && !matches!(
                container.display,
                Some(crate::css::Display::Flex) | Some(crate::css::Display::InlineFlex)
            )
        {
            return false;
        }
        let own = self.get_computed_style(node);
        // **A grid item is stretched to its ROW**, and both axes stretch by
        // default — there is no main/cross split to make, so the block axis is
        // definite whenever the grid's own is. This is what carries a tight
        // constraint through a Flutter `Expanded`, which IS a single-cell grid.
        if grid {
            let align = own
                .align_self
                .or(container.align_items)
                .unwrap_or(crate::css::AlignItems::Stretch);
            return align == crate::css::AlignItems::Stretch
                && self.axis_mode(parent, false) == "declared";
        }
        let along_block_axis =
            if container.flex_direction == Some(crate::css::FlexDirection::Column) {
                // The block axis is the MAIN one: the item's size came from
                // flexing, so it is definite if it grew.
                own.flex_grow.unwrap_or(0.0) > 0.0
            } else {
                // The block axis is the CROSS one, where `stretch` — the
                // initial value of `align-items`, and so of `align-self` —
                // makes the item exactly as tall as the line. That is as
                // definite as a declared length, and it is what lets a key
                // inside a Flutter `Row` resolve `height: 100%`.
                //
                // Any other alignment sizes the item to its own content
                // instead, which is precisely NOT definite.
                let align = own
                    .align_self
                    .or(container.align_items)
                    .unwrap_or(crate::css::AlignItems::Stretch);
                align == crate::css::AlignItems::Stretch
            };
        if !along_block_axis {
            return false;
        }
        self.axis_mode(parent, false) == "declared"
    }

    fn apply_formatting(&mut self, node: NodeId) {
        let Some(display) = self.get_computed_style(node).display else {
            return;
        };
        let spelling = display.as_css().to_string();
        // The INNER context. `inline-flex`/`inline-grid` establish exactly the
        // same one as their block-level counterparts — the `inline-` prefix
        // speaks about the box's OUTER role, which is what `SetChildDisplay`
        // below carries to the parent.
        let context = match display {
            crate::css::Display::Flex | crate::css::Display::InlineFlex => "flex",
            crate::css::Display::Grid | crate::css::Display::InlineGrid => "grid",
            // The table box lays its rows and cells out itself — §17.5. The
            // ROW and ROW GROUP boxes establish nothing: the table reaches
            // through them to the cells, because a row's width is the table's
            // and a column exists across rows that know nothing about it.
            crate::css::Display::Table | crate::css::Display::InlineTable => "table",
            // A row and a row GROUP arrange nothing — the table reaches through
            // them. Left on `normal` they ran block flow over their own cells
            // and stacked what the table had just placed in columns.
            crate::css::Display::TableRow
            | crate::css::Display::TableRowGroup
            | crate::css::Display::TableHeaderGroup
            | crate::css::Display::TableFooterGroup => "table-row",
            _ => "normal",
        };
        self.command(
            node,
            &WidgetCommand::Custom(
                "SetFormatting".into(),
                CommandValue::Text(context.to_string()),
            ),
        );
        // **`flex-direction` has an INITIAL VALUE, and it is `row`** — css-flexbox
        // §5.1. An undeclared property is not an absent one: `display: flex`
        // alone means a ROW, and every browser lays one out that way.
        //
        // Only a DECLARED `flex-direction` ever reached the widget, whose own
        // default is `TopDown` — so a flex container that did not spell the
        // axis out got a column. Nothing declared it wrong; the initial value
        // simply never arrived. A `FlowLayoutPanel`, which is `display: flex;
        // flex-wrap: wrap` and nothing else, stacked its buttons vertically at
        // full width instead of flowing them across and wrapping.
        //
        // Sent from here because this is where the box BECOMES a flex
        // container, which is the moment the initial value starts to apply.
        if context == "flex" {
            let direction = self
                .get_computed_style(node)
                .flex_direction
                .map(|d| d.as_css().to_string())
                .unwrap_or_else(|| "row".to_string());
            self.command(
                node,
                &WidgetCommand::Custom("SetFlexDirection".into(), CommandValue::Text(direction)),
            );
        }
        // **Whether its own block size is definite**, which is the other half
        // of what a formatting context needs before it can size anything: with
        // `height: auto` there is no main size for a column to divide, so its
        // items resolve against their content instead.
        //
        // Here rather than in `announce_child_display`, which is the natural
        // reading and does not work: that runs BEFORE the widget is linked, so
        // it addresses its container by name and a send to the box itself would
        // reach nothing. This is the same send as in `apply_style_property`,
        // for the box that never takes a `height` write at all.
        let own = self.axis_mode(node, false).to_string();
        self.command(
            node,
            &WidgetCommand::Custom("SetHeightMode".into(), CommandValue::Text(own)),
        );
        self.tell_container(node, "SetChildDisplay", &spelling);
    }

    /// Tell the container where this item sits — both axes, whole, on any
    /// change, for the same reason the border is sent whole.
    fn send_grid_placement(&mut self, node: NodeId) {
        let props = self.get_computed_style(node);
        let line = |l: Option<crate::css::GridLine>| {
            l.unwrap_or(crate::css::GridLine::Auto).as_css()
        };
        let spec = format!(
            "{},{},{},{}",
            line(props.grid_column_start),
            line(props.grid_column_end),
            line(props.grid_row_start),
            line(props.grid_row_end),
        );
        self.tell_container(node, "SetChildGridArea", &spec);
    }

    /// Hand the widget its whole border: four widths, a style and a colour.
    ///
    /// **Style decides whether anything paints at all.** CSS's initial
    /// `border-style` is `none`, and a `none` border is zero-width however wide
    /// it was declared — which is why a `border-width` on its own correctly
    /// draws nothing, and why this reads the style rather than inferring from
    /// the width.
    ///
    /// One colour and one style for the box, per-side widths. Per-side colours
    /// and styles are legal CSS and rare; stated as the boundary rather than
    /// half-supported.
    fn send_border(&mut self, node: NodeId) {
        let props = self.get_computed_style(node);
        let (widths, styles, colours) = (props.border_width, props.border_style, props.border_color);
        // **Each side decides for itself.** A side whose style is `none` is
        // zero-wide however wide it was declared — CSS's rule — so the style is
        // what is asked, per side, rather than one style for the box. That is
        // what `border-bottom: 1px solid` with no other border means, and it is
        // the ordinary way to draw a rule under a heading.
        let side = |w: Option<f32>, st: Option<crate::css::BorderStyle>| -> f32 {
            match st.unwrap_or(crate::css::BorderStyle::None) {
                crate::css::BorderStyle::None => 0.0,
                _ => w.unwrap_or(0.0).max(0.0),
            }
        };
        let spec = format!(
            "{},{},{},{};{:#010x},{:#010x},{:#010x},{:#010x}",
            side(widths.top, styles.top),
            side(widths.right, styles.right),
            side(widths.bottom, styles.bottom),
            side(widths.left, styles.left),
            colours.top.unwrap_or(0xFF00_0000),
            colours.right.unwrap_or(0xFF00_0000),
            colours.bottom.unwrap_or(0xFF00_0000),
            colours.left.unwrap_or(0xFF00_0000),
        );
        self.command(
            node,
            &WidgetCommand::Custom("SetBorderBox".into(), CommandValue::Text(spec)),
        );
    }

    /// One geometry axis as the CASCADE computed it, spelled the way the
    /// declaration store would spell it so both routes reach the same parser.
    ///
    /// `None` when nothing in the cascade set that axis — which is the state
    /// that means "the container decides", not zero.
    fn computed_axis(&self, node: NodeId, axis: &str) -> Option<String> {
        let props = self.get_computed_style(node);
        let length = match axis {
            "left" => props.offsets.left,
            "top" => props.offsets.top,
            "width" => props.width,
            "height" => props.height,
            _ => None,
        }?;
        Some(length.to_string())
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
                let value = self.style(node).map(|s| s.get(axis).to_string());
                // The declaration store first — it is what the program wrote —
                // and the CASCADE when the program wrote nothing. Without the
                // second half a box positioned entirely by a stylesheet had its
                // `position` honoured and its coordinates lost, which lands it
                // out of flow at the origin: worse than either answer alone.
                let value = match value {
                    Some(v) if !v.trim().is_empty() => v,
                    _ => self.computed_axis(node, axis)?,
                };
                self.px(node, axis, &value)
                    .map(|px| ((*axis).to_string(), px))
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

    /// Re-run the layout of the TABLE this node is inside, if any.
    ///
    /// **A row and a row group establish no formatting context**, so re-laying
    /// out a cell's parent does nothing at all — the box that decides where a
    /// cell goes is the table, several levels up. Every other container
    /// arranges its own children, which is why `relayout_parent` has been
    /// enough everywhere else.
    ///
    /// Without this, a cell appended to a row that had ALREADY joined its table
    /// was never placed: the table's last layout ran before the cell existed
    /// and nothing asked it to look again. The tell was exact — every row laid
    /// out except the LAST one, in every table, because the last row is the only
    /// one no later insertion happens to re-trigger.
    ///
    /// ⚠ The walk stops at the FIRST enclosing table. A table nested in a cell
    /// re-lays itself out, but its new size does not propagate to the outer
    /// table's column widths until something else touches the outer one.
    fn relayout_enclosing_table(&mut self, node: NodeId) {
        use crate::css::Display;
        let mut at = node;
        // A cell is inside a row inside an optional group inside the table, so
        // the walk is short — but it is a walk, not a fixed number of steps.
        while let Some(parent) = self.nodes.get(&at).and_then(|n| n.parent) {
            if parent == DOCUMENT {
                return;
            }
            if matches!(
                self.get_computed_style(parent).display,
                Some(Display::Table) | Some(Display::InlineTable)
            ) {
                // **Re-setting the SAME rect is the invalidation, not a
                // no-op.** `set_rect` re-runs `relayout()`, which is what
                // re-reads the rows, rebuilds the cell grid and re-measures
                // the columns — the table's geometry is unchanged, but what
                // fills it is not. ⚠ Anyone adding an "unchanged rect → skip"
                // fast path to `set_rect` silently kills table layout.
                if let Some(w) = self.widget_mut(parent) {
                    let unchanged = w.rect();
                    w.set_rect(unchanged);
                }
                // **And the table's own BOX has to follow its rows.** Its
                // height is its content's — §17.5.3 — so a table that gained a
                // row is taller and one that lost content is shorter, and the
                // box it draws is the only thing that says where its border
                // goes. Without this the internals were re-laid out inside a
                // rectangle from the previous pass: a table whose last row
                // arrived after the box was measured had that row CLIPPED by
                // its own bottom border, and one that was measured while it
                // still had more content kept the taller box and drew a large
                // empty area under its last row. Both were visible in a
                // capture and in neither case was a single cell misplaced.
                self.size_to_content(parent);
                return;
            }
            at = parent;
        }
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
        // The BODY is the root of the flow, so a change under it re-runs the
        // root pass — the same answer `DOCUMENT` gives, because appending to
        // the document IS appending to the body. Without this, content moved
        // into a real body stopped triggering the pass that resolves
        // percentages and sizes flex rows, and only its widget rect was
        // re-set: a `width: 100%` button kept its 96px default.
        if parent == DOCUMENT || Some(parent) == self.body {
            self.relayout_body();
            return;
        }
        if let Some(w) = self.widget_mut(parent) {
            let rect = w.rect();
            w.set_rect(rect);
        }
    }

    /// **The body is a block container** — stack its in-flow children.
    ///
    /// `<body>` is `display: block` and block-level children flow down it. This
    /// did nothing at all: `insert_widget` placed every top-level element at
    /// the rect it happened to carry, and `relayout_parent` returned early for
    /// `DOCUMENT` with a comment saying the body positions absolutely. For a
    /// frontend that emits coordinates that is true and invisible — every VCL
    /// control carries `left`/`top`, so every one is out of flow. For a page
    /// that appends a `<div>` and three `<button>`s with no coordinates, all
    /// four landed at (0, 0) and only the last painted.
    ///
    /// Out-of-flow children are skipped, which is what makes both work at once:
    /// a positioned element keeps its own coordinates and takes no space, and
    /// only what is genuinely in flow stacks. That is CSS's rule, not a
    /// special case for the body.
    fn relayout_body(&mut self) {
        let origin = self.containing_block(DOCUMENT);
        // The body's children, not the document's — the document has exactly
        // one, `<html>`. Walking the document here meant that the moment tree
        // construction put a real body in the way, this pass stopped reaching
        // any content at all: percentages never resolved (a `width: 100%`
        // button kept its 96px default) and nothing was placed.
        let root = self.body.unwrap_or(DOCUMENT);
        for child in self.child_nodes(root) {
            // Only elements are boxes. A text node is a run of the body's own
            // line; a comment, a CDATA section and a processing instruction
            // are not even that; and metadata renders nothing and must not
            // occupy a row of blank space.
            if !self.is_element(child) {
                continue;
            }
            let skip = self
                .node(child)
                .map(|n| n.dock.is_some() || renders_nothing(&n.tag, &n.input_type))
                .unwrap_or(true);
            if skip || self.positions_itself(child) {
                continue;
            }
            // **The two halves of sizing, in the only order they work in.** A
            // block's width comes from the container and its height comes from
            // its content, so the width has to be settled before the content is
            // measured — the width is what decides where the text wrapped, and
            // the wrapping is what decides how many line boxes there are.
            //
            // This is ALL that is left of this pass. Placement — margins,
            // `position: relative` offsets, the running `y` — belonged here
            // only while the root had no box to do it; `layout_normal_flow`
            // already handles all three (`relative_offset` in `flow_layout.rs`)
            // and doing it twice would fight it.
            self.size_subtree(child, origin.w);
        }
        // **The body lays out its own children.** This used to place them by
        // hand — `x = origin.x`, then `y += height` per child — which is a
        // block stacker and ONLY a block stacker: it never consulted `display`,
        // so two inline-level boxes at the root could not share a line however
        // the cascade described them. Every other container in the document
        // runs `layout_normal_flow`, which does line boxes, inline-level
        // sharing and margins; the root was the one box that did not, purely
        // because it had no element to be.
        //
        // Now it has one. Re-setting the rect is how `relayout_parent` re-runs
        // any other container, and it is the whole of what the root needs — the
        // sizing pass above still settles widths and percentages first, because
        // a block's width has to be known before its content can be measured.
        if let Some(body) = self.body {
            if let Some(w) = self.widget_mut(body) {
                let rect = w.rect();
                w.set_rect(rect);
            }
        }
    }

    /// **`width: auto` on an in-flow block fills its containing block** —
    /// CSS 2.1 §10.3.3.
    ///
    /// The rule that makes a paragraph as wide as the page rather than as wide
    /// as a guess, and the reason the height half of intrinsic sizing is
    /// answerable at all: the width is known BEFORE the content is measured, so
    /// only the height has to come back from it.
    ///
    /// It is the MARGIN box that fills, which is why the margins come off the
    /// width rather than being added to it. `box-sizing` does not enter: it
    /// decides how a *declared* length maps onto the border box, and this box
    /// has not declared one.
    ///
    /// Only a box the cascade calls `block`. Not "anything without a declared
    /// width" — an `inline-block` is shrink-to-fit and a box with no `display`
    /// at all is one this stylesheet has never met, and stretching either
    /// across the page would be a guess wearing a rule's clothes.
    fn fill_available_width(&mut self, node: NodeId, available: f32) -> bool {
        let props = self.get_computed_style(node);
        // **Block-LEVEL, not `display: block`.** `display` says two things, and
        // this is the outer half: a `flex` or `grid` box participates in its
        // parent's flow exactly as a block does, so `width: auto` fills the
        // containing block for all three. Testing for `Block` alone starved
        // every flex and grid container to `default_size`'s 200px guess — and
        // since Flutter's `Scaffold`, `Column` and `Row` are all
        // `div;display:flex`, the entire app rendered as a 200px strip.
        //
        // The `inline-*` forms are excluded because their outer display really
        // is inline-level: they are shrink-to-fit, and stretching one would be
        // the same guess this rule exists to remove.
        let block_level = matches!(
            props.display,
            Some(crate::css::Display::Block)
                | Some(crate::css::Display::Flex)
                | Some(crate::css::Display::Grid)
        );
        if props.width.is_some() || !block_level {
            return false;
        }
        // **In a flex or grid container the CONTAINER sizes the item** — §7.1
        // and §11.1. "Fill the containing block" is normal flow's rule for a
        // block box, and applying it to a flex item overwrote the width the
        // flex algorithm had just handed out: every cell in a Flutter row was
        // stretched back to the full row, so a `width: 100%` child of one
        // resolved against the whole row and painted its label into the next
        // column.
        if self
            .nodes
            .get(&node)
            .and_then(|n| n.parent)
            .map(|p| self.establishes_independent_track_layout(p))
            .unwrap_or(false)
        {
            return false;
        }
        let width = available - self.box_edges(node).margin.horizontal();
        if width <= 0.0 {
            return false;
        }
        let Some(w) = self.widget_mut(node) else {
            return false;
        };
        let rect = w.rect();
        if (rect.w - width).abs() < 0.5 {
            return false;
        }
        w.set_rect(LayoutRect::new(rect.x, rect.y, width, rect.h));
        true
    }

    /// Has this element been positioned by its own coordinates?
    ///
    /// The bridge until every frontend declares `position` explicitly. In CSS
    /// `left`/`top` do nothing on a `position: static` box, so a frontend
    /// setting them and expecting them to count is *already* saying
    /// `position: absolute` — it simply has not said so yet. Reading the
    /// declaration store means this answers from what the program actually set,
    /// not from a guess.
    pub(crate) fn positions_itself(&self, node: NodeId) -> bool {
        // **The CASCADE decides, not the element's own declarations.** This
        // read the declaration store alone, so `position` could only ever be
        // honoured when the program wrote it directly onto the element: a
        // `<style>` rule — or a UA rule — computes correctly and then had no
        // effect at all, because author-sheet results live in `computed` and
        // this never looked there. Measured, not assumed: a `#box { position:
        // absolute; left: 40px }` rule computed `Absolute` and laid out at
        // (0, 0).
        //
        // The declaration store is still consulted underneath, because it is
        // the CSSOM's own record and carries the `left`/`top` inference below,
        // which has no computed equivalent.
        use crate::css::Position;
        match self.get_computed_style(node).position {
            Some(Position::Absolute) | Some(Position::Fixed) => return true,
            Some(Position::Static) | Some(Position::Relative) | Some(Position::Sticky) => {
                return false;
            }
            None => {}
        }
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
    /// `node.parentElement` — the parent, but only when it IS an element.
    pub fn parent_element(&self, node: NodeId) -> Option<NodeId> {
        let parent = self.parent_node(node)?;
        self.is_element(parent).then_some(parent)
    }

    /// `node.firstChild` / `lastChild` — of ANY kind.
    pub fn first_child(&self, node: NodeId) -> Option<NodeId> {
        self.child_nodes(node).first().copied()
    }

    pub fn last_child(&self, node: NodeId) -> Option<NodeId> {
        self.child_nodes(node).last().copied()
    }

    pub fn has_child_nodes(&self, node: NodeId) -> bool {
        !self.child_nodes(node).is_empty()
    }

    /// `node.contains(other)` — DOM §4.4. A node CONTAINS ITSELF.
    pub fn contains(&self, node: NodeId, other: NodeId) -> bool {
        let mut current = Some(other);
        while let Some(id) = current {
            if id == node {
                return true;
            }
            current = self.parent_node(id);
        }
        false
    }

    /// `element.children` — ELEMENT children only, unlike `childNodes`.
    pub fn children(&self, node: NodeId) -> Vec<NodeId> {
        self.child_nodes(node)
            .into_iter()
            .filter(|c| self.is_element(*c))
            .collect()
    }

    pub fn first_element_child(&self, node: NodeId) -> Option<NodeId> {
        self.children(node).first().copied()
    }

    pub fn last_element_child(&self, node: NodeId) -> Option<NodeId> {
        self.children(node).last().copied()
    }

    pub fn child_element_count(&self, node: NodeId) -> usize {
        self.children(node).len()
    }

    /// `element.nextElementSibling` — skips text and comments.
    pub fn next_element_sibling(&self, node: NodeId) -> Option<NodeId> {
        let parent = self.parent_node(node)?;
        let siblings = self.children(parent);
        let at = siblings.iter().position(|n| *n == node)?;
        siblings.get(at + 1).copied()
    }

    pub fn previous_element_sibling(&self, node: NodeId) -> Option<NodeId> {
        let parent = self.parent_node(node)?;
        let siblings = self.children(parent);
        let at = siblings.iter().position(|n| *n == node)?;
        if at == 0 { None } else { siblings.get(at - 1).copied() }
    }

    /// `element.matches(selectors)`.
    ///
    /// Answered through `querySelectorAll`, so one selector engine decides
    /// both questions — a `matches` that disagreed with the selector that
    /// found the element would be worse than not having it.
    pub fn matches(&self, node: NodeId, selectors: &str) -> bool {
        self.query_selector_all(selectors).contains(&node)
    }

    /// `element.closest(selectors)` — nearest ancestor-OR-SELF that matches.
    pub fn closest(&self, node: NodeId, selectors: &str) -> Option<NodeId> {
        let candidates = self.query_selector_all(selectors);
        let mut current = Some(node);
        while let Some(id) = current {
            if candidates.contains(&id) {
                return Some(id);
            }
            current = self.parent_node(id);
        }
        None
    }

    /// `getElementsByClassName(names)` — every element carrying ALL the named
    /// classes, in tree order.
    pub fn get_elements_by_class_name(&self, names: &str) -> Vec<NodeId> {
        let wanted: Vec<&str> = names.split_whitespace().collect();
        if wanted.is_empty() {
            return Vec::new();
        }
        self.get_elements_by_tag_name("*")
            .into_iter()
            .filter(|id| wanted.iter().all(|c| self.class_list_contains(*id, c)))
            .collect()
    }

    /// `element.hasAttributes()`.
    ///
    /// `&mut` because `getAttributeNames` reconstructs `style`/`value`/`checked`
    /// from the live control, which this engine has to ask the widget for.
    pub fn has_attributes(&mut self, node: NodeId) -> bool {
        !self.get_attribute_names(node).is_empty()
    }

    /// `element.toggleAttribute(name)` — answers its presence AFTERWARDS.
    pub fn toggle_attribute(&mut self, node: NodeId, name: &str) -> bool {
        if self.has_attribute(node, name) {
            self.remove_attribute(node, name);
            false
        } else {
            self.set_attribute(node, name, "");
            true
        }
    }

    /// `element.className` / `element.id`.
    pub fn class_name(&mut self, node: NodeId) -> String {
        self.get_attribute(node, "class").unwrap_or_default()
    }

    pub fn set_class_name(&mut self, node: NodeId, value: &str) {
        self.set_attribute(node, "class", value);
    }

    pub fn id(&mut self, node: NodeId) -> String {
        self.get_attribute(node, "id").unwrap_or_default()
    }

    pub fn set_id(&mut self, node: NodeId, value: &str) {
        self.set_attribute(node, "id", value);
    }

    /// `document.documentElement` / `head`.
    pub fn document_element(&self) -> Option<NodeId> {
        self.query_selector("html").or_else(|| self.body())
    }

    pub fn head(&self) -> Option<NodeId> {
        self.query_selector("head")
    }

    // ─── Node comparison and normalisation (DOM §4.4, §4.5) ─────────────────

    /// `node.normalize()` — drop empty text nodes and merge adjacent ones.
    ///
    /// What makes a tree built by repeated `appendChild(createTextNode(..))`
    /// compare equal to the same tree built by a parser. Until it runs, two
    /// documents that render identically can disagree under
    /// [`is_equal_node`](Self::is_equal_node) purely over where the text was
    /// split, which is a difference no reader can see.
    pub fn normalize(&mut self, node: NodeId) {
        let children = self.child_nodes(node);
        for &child in &children {
            self.normalize(child);
        }
        // The run being merged INTO. Reset by any non-text child, because two
        // text nodes with an element between them are not adjacent.
        let mut run: Option<NodeId> = None;
        for child in children {
            if self.node(child).map(|n| n.kind) != Some(NodeKind::Text) {
                run = None;
                continue;
            }
            let data = self.node(child).map(|n| n.data.clone()).unwrap_or_default();
            if data.is_empty() {
                // A zero-length text node is removed outright — and does NOT
                // break the run, so `a`, ``, `b` merges to `ab`.
                self.remove_child(node, child);
                continue;
            }
            match run {
                Some(first) => {
                    if let Some(n) = self.nodes.get_mut(&first) {
                        n.data.push_str(&data);
                    }
                    self.remove_child(node, child);
                }
                None => run = Some(child),
            }
        }
    }

    /// `node.isEqualNode(other)` — structural equality, not identity.
    ///
    /// Attribute ORDER is deliberately not part of it (DOM §4.4): an attribute
    /// list is a set as far as equality is concerned, so two elements written
    /// `<a id x>` and `<a x id>` are equal.
    pub fn is_equal_node(&self, a: NodeId, b: NodeId) -> bool {
        let (Some(x), Some(y)) = (self.node(a), self.node(b)) else {
            // The document node itself has no `DomNode`, so identity is the
            // only thing that can answer for it — but ONLY for it. Two ids
            // that name nothing are not nodes, and two non-nodes are not
            // equal to each other.
            return a == b && a == DOCUMENT;
        };
        if x.kind != y.kind {
            return false;
        }
        match x.kind {
            NodeKind::Element => {
                if x.tag != y.tag || x.namespace != y.namespace {
                    return false;
                }
                if x.attributes.len() != y.attributes.len() {
                    return false;
                }
                let matched = x.attributes.iter().all(|a| {
                    y.attributes.iter().any(|b| {
                        a.namespace == b.namespace && a.name == b.name && a.value == b.value
                    })
                });
                if !matched {
                    return false;
                }
            }
            _ => {
                if x.data != y.data {
                    return false;
                }
            }
        }
        x.children.len() == y.children.len()
            && x.children
                .iter()
                .zip(y.children.iter())
                .all(|(&c, &d)| self.is_equal_node(c, d))
    }

    /// The root-to-node path, used to put two nodes in tree order.
    fn ancestor_path(&self, node: NodeId) -> Vec<NodeId> {
        let mut path = vec![node];
        let mut current = node;
        while let Some(parent) = self.parent_node(current) {
            path.push(parent);
            current = parent;
        }
        path.reverse();
        path
    }

    /// `node.compareDocumentPosition(other)` — DOM §4.4's bitmask.
    ///
    /// `DISCONNECTED 0x01`, `PRECEDING 0x02`, `FOLLOWING 0x04`,
    /// `CONTAINS 0x08`, `CONTAINED_BY 0x10`, `IMPLEMENTATION_SPECIFIC 0x20`.
    /// Containment always arrives paired with a direction, which is why the
    /// answers below are two bits and not one.
    pub fn compare_document_position(&self, a: NodeId, b: NodeId) -> u16 {
        const DISCONNECTED: u16 = 0x01;
        const PRECEDING: u16 = 0x02;
        const FOLLOWING: u16 = 0x04;
        const CONTAINS: u16 = 0x08;
        const CONTAINED_BY: u16 = 0x10;
        const IMPLEMENTATION_SPECIFIC: u16 = 0x20;

        if a == b {
            return 0;
        }
        let pa = self.ancestor_path(a);
        let pb = self.ancestor_path(b);
        if pa.first() != pb.first() {
            // The spec lets an implementation pick the direction for two
            // detached trees, but requires it to be CONSISTENT. Ordering by
            // the node handle is the cheapest thing that is.
            let direction = if a < b { FOLLOWING } else { PRECEDING };
            return DISCONNECTED | IMPLEMENTATION_SPECIFIC | direction;
        }
        // The first index at which the two paths part company.
        let split = pa
            .iter()
            .zip(pb.iter())
            .position(|(x, y)| x != y)
            .unwrap_or_else(|| pa.len().min(pb.len()));
        if split == pa.len() {
            return CONTAINED_BY | FOLLOWING;
        }
        if split == pb.len() {
            return CONTAINS | PRECEDING;
        }
        // Same parent at `split - 1`; their order there is the answer.
        let siblings = self.child_nodes(pa[split - 1]);
        let ia = siblings.iter().position(|n| *n == pa[split]);
        let ib = siblings.iter().position(|n| *n == pb[split]);
        match (ia, ib) {
            (Some(ia), Some(ib)) if ia < ib => FOLLOWING,
            (Some(_), Some(_)) => PRECEDING,
            _ => DISCONNECTED | IMPLEMENTATION_SPECIFIC | FOLLOWING,
        }
    }

    // ─── The rest of ParentNode / ChildNode (DOM §4.2.6) ────────────────────

    /// `parent.append(...nodes)`.
    pub fn append(&mut self, parent: NodeId, nodes: &[NodeId]) {
        for &node in nodes {
            self.append_child(parent, node);
        }
    }

    /// `parent.prepend(...nodes)`.
    ///
    /// Every node goes before the ORIGINAL first child, not before whatever is
    /// first at the time — inserting each before the running first child would
    /// hand the caller its nodes back reversed.
    pub fn prepend(&mut self, parent: NodeId, nodes: &[NodeId]) {
        let reference = self.first_child(parent);
        for &node in nodes {
            self.insert_before(parent, node, reference);
        }
    }

    /// `node.before(...nodes)`.
    pub fn before(&mut self, node: NodeId, nodes: &[NodeId]) {
        let Some(parent) = self.parent_node(node) else { return };
        for &new in nodes {
            self.insert_before(parent, new, Some(node));
        }
    }

    /// `node.after(...nodes)`.
    pub fn after(&mut self, node: NodeId, nodes: &[NodeId]) {
        let Some(parent) = self.parent_node(node) else { return };
        let reference = self.next_sibling(node);
        for &new in nodes {
            self.insert_before(parent, new, reference);
        }
    }

    /// `node.replaceWith(...nodes)`.
    pub fn replace_with(&mut self, node: NodeId, nodes: &[NodeId]) {
        let Some(parent) = self.parent_node(node) else { return };
        for &new in nodes {
            self.insert_before(parent, new, Some(node));
        }
        self.remove_child(parent, node);
    }

    /// `element.insertAdjacentElement(position, element)`.
    ///
    /// Returns the inserted element, or `None` for an unknown position or a
    /// placement with no parent to hold it — the IDL's `null`.
    pub fn insert_adjacent_element(
        &mut self,
        node: NodeId,
        position: &str,
        other: NodeId,
    ) -> Option<NodeId> {
        let placed = match position.to_ascii_lowercase().as_str() {
            "beforebegin" => {
                let parent = self.parent_node(node)?;
                self.insert_before(parent, other, Some(node))
            }
            "afterbegin" => {
                let reference = self.first_child(node);
                self.insert_before(node, other, reference)
            }
            "beforeend" => self.append_child(node, other),
            "afterend" => {
                let parent = self.parent_node(node)?;
                let reference = self.next_sibling(node);
                self.insert_before(parent, other, reference)
            }
            _ => false,
        };
        placed.then_some(other)
    }

    /// `element.insertAdjacentText(position, data)`.
    pub fn insert_adjacent_text(&mut self, node: NodeId, position: &str, data: &str) {
        let text = self.create_text_node(data);
        self.insert_adjacent_element(node, position, text);
    }

    // ─── HTMLElement: dataset, innerText, tabIndex, click (HTML §3.2.6) ─────

    /// `element.dataset` — every `data-*` attribute, keyed the way the IDL
    /// names it rather than the way it is spelled in the markup.
    ///
    /// A snapshot, not the live `DOMStringMap`: Rust has nowhere to put an
    /// object whose property writes reach back into the element, so the pair
    /// [`set_dataset`](Self::set_dataset) / [`remove_dataset`](Self::remove_dataset)
    /// carries the write side.
    /// Sorted by key, which is the ONLY order the other engine can also
    /// promise — it keeps its attributes in a map and has no document order
    /// left to report. Two engines that answer in different orders would make
    /// `dataset(n)[0]` mean different things, so neither reports its own.
    pub fn dataset(&self, node: NodeId) -> Vec<(String, String)> {
        let mut pairs: Vec<(String, String)> = self
            .node(node)
            .map(|n| {
                n.attributes
                    .iter()
                    .filter_map(|a| dataset_key(&a.name).map(|k| (k, a.value.clone())))
                    .collect()
            })
            .unwrap_or_default();
        pairs.sort();
        pairs
    }

    /// `element.dataset[key]`.
    pub fn dataset_get(&self, node: NodeId, key: &str) -> Option<String> {
        let name = dataset_attribute(key);
        self.node(node)
            .and_then(|n| n.attributes.iter().find(|a| a.name == name))
            .map(|a| a.value.clone())
    }

    /// `element.dataset[key] = value`.
    pub fn set_dataset(&mut self, node: NodeId, key: &str, value: &str) {
        let name = dataset_attribute(key);
        self.set_attribute(node, &name, value);
    }

    /// `delete element.dataset[key]`.
    pub fn remove_dataset(&mut self, node: NodeId, key: &str) {
        let name = dataset_attribute(key);
        self.remove_attribute(node, &name);
    }

    /// `element.innerText` — the RENDERED text.
    ///
    /// The difference from `textContent` is the cascade: a subtree the styles
    /// hid contributes nothing, and a block boundary becomes a line break.
    /// Reading `textContent` where a page asked for `innerText` is how hidden
    /// markup leaks into a string a user is shown.
    pub fn inner_text(&mut self, node: NodeId) -> String {
        let mut out = String::new();
        self.inner_text_into(node, &mut out);
        out.trim_matches('\n').to_string()
    }

    fn inner_text_into(&mut self, node: NodeId, out: &mut String) {
        match self.node(node).map(|n| n.kind) {
            Some(NodeKind::Text) => {
                if let Some(n) = self.node(node) {
                    let data = n.data.clone();
                    out.push_str(&data);
                }
            }
            Some(NodeKind::Element) => {
                let display = self.computed_style_property(node, "display");
                if display == "none" {
                    return;
                }
                let block = display != "inline";
                if block && !out.is_empty() && !out.ends_with('\n') {
                    out.push('\n');
                }
                for child in self.child_nodes(node) {
                    self.inner_text_into(child, out);
                }
                if block && !out.is_empty() && !out.ends_with('\n') {
                    out.push('\n');
                }
            }
            _ => {}
        }
    }

    /// `element.tabIndex`.
    ///
    /// The default is not zero for everything: HTML §6.6.3 gives 0 to what is
    /// focusable without the attribute and −1 to the rest, so a `<div>` and a
    /// `<button>` answer differently with no `tabindex` on either.
    pub fn tab_index(&self, node: NodeId) -> i32 {
        let declared = self
            .node(node)
            .and_then(|n| n.attributes.iter().find(|a| a.name == "tabindex"))
            .and_then(|a| a.value.trim().parse::<i32>().ok());
        if let Some(value) = declared {
            return value;
        }
        let Some(n) = self.node(node) else { return -1 };
        match n.tag.as_str() {
            // A link is only in the sequence once it has somewhere to go.
            "a" | "area" => {
                if n.attributes.iter().any(|a| a.name == "href") {
                    0
                } else {
                    -1
                }
            }
            "button" | "input" | "select" | "textarea" | "iframe" => 0,
            _ => -1,
        }
    }

    /// `element.tabIndex = index`.
    pub fn set_tab_index(&mut self, node: NodeId, index: i32) {
        self.set_attribute(node, "tabindex", &index.to_string());
    }

    /// `element.click()` — a synthetic click, dispatched like a real one.
    ///
    /// It goes through [`dispatch_event`](Self::dispatch_event) rather than
    /// calling handlers directly so that capture, bubbling and
    /// `stopPropagation` behave exactly as they do for a user's click. A page
    /// cannot tell the two apart, which is the point.
    pub fn click(&mut self, node: NodeId) {
        let mut event = DomEvent::new(node, "click");
        self.dispatch_event(&mut event);
    }

    // ─── CSSOM-View (§5) ───────────────────────────────────────────────────

    /// `element.getClientRects()`.
    ///
    /// One rect per box the element generates. A form's inline runs carry no
    /// geometry of their own, so a wrapped inline reports the single border
    /// box here where a line-box engine would report one rect per line.
    pub fn get_client_rects(&mut self, node: NodeId) -> Vec<crate::layout::LayoutRect> {
        self.get_bounding_client_rect(node).into_iter().collect()
    }

    /// `element.scrollIntoView()`.
    ///
    /// A form's document has no scrollable viewport — every box is laid out
    /// inside the window it was given and nothing is clipped out of view, so
    /// the element is already in view and there is no scroll offset to move.
    /// Present because a page calls it unconditionally; it is not standing in
    /// for scrolling that ought to be happening.
    pub fn scroll_into_view(&mut self, _node: NodeId) {}

    /// `element.offsetParent` — the nearest POSITIONED ancestor, else the body.
    pub fn offset_parent(&mut self, node: NodeId) -> Option<NodeId> {
        let mut current = self.parent_node(node)?;
        loop {
            let tag = self.node(current).map(|n| n.tag.clone()).unwrap_or_default();
            if tag == "body" || tag == "html" {
                return self.body();
            }
            if self.computed_style_property(current, "position") != "static" {
                return Some(current);
            }
            current = self.parent_node(current)?;
        }
    }

    /// `element.offsetTop` — the border box's top edge, relative to
    /// [`offset_parent`](Self::offset_parent) rather than to the viewport.
    pub fn offset_top(&mut self, node: NodeId) -> f32 {
        let own = self.get_bounding_client_rect(node).map(|r| r.y).unwrap_or(0.0);
        let origin = self
            .offset_parent(node)
            .and_then(|p| self.get_bounding_client_rect(p))
            .map(|r| r.y)
            .unwrap_or(0.0);
        own - origin
    }

    /// `element.offsetLeft` — see [`offset_top`](Self::offset_top).
    pub fn offset_left(&mut self, node: NodeId) -> f32 {
        let own = self.get_bounding_client_rect(node).map(|r| r.x).unwrap_or(0.0);
        let origin = self
            .offset_parent(node)
            .and_then(|p| self.get_bounding_client_rect(p))
            .map(|r| r.x)
            .unwrap_or(0.0);
        own - origin
    }

    // ─── The namespaced accessors that were missing their pair ─────────────

    /// `element.hasAttributeNS(namespace, localName)`.
    pub fn has_attribute_ns(&self, node: NodeId, namespace: &str, local_name: &str) -> bool {
        let namespace = (!namespace.is_empty()).then_some(namespace);
        self.node(node)
            .and_then(|n| n.attribute_ns(namespace, local_name))
            .is_some()
    }

    /// `element.removeAttributeNS(namespace, localName)`.
    pub fn remove_attribute_ns(&mut self, node: NodeId, namespace: &str, local_name: &str) {
        let namespace = (!namespace.is_empty()).then(|| namespace.to_string());
        if let Some(n) = self.nodes.get_mut(&node) {
            n.attributes
                .retain(|a| !(a.namespace == namespace && a.local_name() == local_name));
        }
    }

    /// `element.hasAttribute(name)`.
    ///
    /// Not `get_attribute(..).is_some()` at the call site: an attribute present
    /// with an EMPTY value is present, and `checked=""` is exactly that.
    pub fn has_attribute(&self, node: NodeId, name: &str) -> bool {
        let name = self.fold_name(name);
        self.nodes
            .get(&node)
            .map(|n| n.attributes.iter().any(|a| a.name == name))
            .unwrap_or(false)
    }

    /// `element.tagName`.
    pub fn tag_name(&self, node: NodeId) -> Option<&str> {
        self.nodes.get(&node).map(|n| n.tag.as_str())
    }

    /// `element.offsetWidth` / `offsetHeight` — the laid-out border box, as an
    /// integer count of pixels. Zero for a node with no box, which is what a
    /// browser answers for `display: none`.
    pub fn offset_width(&mut self, node: NodeId) -> f32 {
        self.get_bounding_client_rect(node).map(|r| r.w).unwrap_or(0.0)
    }

    pub fn offset_height(&mut self, node: NodeId) -> f32 {
        self.get_bounding_client_rect(node).map(|r| r.h).unwrap_or(0.0)
    }

    /// `node.nextSibling` / `previousSibling` — the next node of ANY kind, not
    /// just the next element. `None` at the end of the list.
    pub fn next_sibling(&self, node: NodeId) -> Option<NodeId> {
        let parent = self.parent_node(node)?;
        let siblings = self.child_nodes(parent);
        let i = siblings.iter().position(|n| *n == node)?;
        siblings.get(i + 1).copied()
    }

    pub fn previous_sibling(&self, node: NodeId) -> Option<NodeId> {
        let parent = self.parent_node(node)?;
        let siblings = self.child_nodes(parent);
        let i = siblings.iter().position(|n| *n == node)?;
        if i == 0 { None } else { siblings.get(i - 1).copied() }
    }

    // ── `element.classList` ──────────────────────────────────────────────
    //
    // Flattened rather than returning a `DOMTokenList`: the token list is a
    // LIVE object over the `class` attribute, and there is no object to hand
    // back across this surface. The operations are the list's, and each one
    // reads and rewrites the attribute, so the attribute stays the single
    // source of truth — `getAttribute("class")` and `classList` cannot
    // disagree.

    fn class_tokens(&self, node: NodeId) -> Vec<String> {
        self.nodes
            .get(&node)
            .and_then(|n| n.attributes.iter().find(|a| a.name == "class"))
            .map(|a| a.value.split_whitespace().map(str::to_string).collect())
            .unwrap_or_default()
    }

    fn write_class_tokens(&mut self, node: NodeId, tokens: &[String]) {
        self.set_attribute(node, "class", &tokens.join(" "));
    }

    /// `classList.add(token)`. Adding a token already present does nothing —
    /// the list is a SET, so this must not produce `"a a"`.
    pub fn class_list_add(&mut self, node: NodeId, class: &str) {
        let mut tokens = self.class_tokens(node);
        if tokens.iter().any(|t| t == class) {
            return;
        }
        tokens.push(class.to_string());
        self.write_class_tokens(node, &tokens);
    }

    /// `classList.remove(token)` — every occurrence, since a malformed
    /// attribute may carry duplicates.
    pub fn class_list_remove(&mut self, node: NodeId, class: &str) {
        let mut tokens = self.class_tokens(node);
        let before = tokens.len();
        tokens.retain(|t| t != class);
        if tokens.len() != before {
            self.write_class_tokens(node, &tokens);
        }
    }

    /// `classList.toggle(token)` — returns whether the token is present
    /// AFTERWARDS, which is what the IDL returns.
    pub fn class_list_toggle(&mut self, node: NodeId, class: &str) -> bool {
        if self.class_list_contains(node, class) {
            self.class_list_remove(node, class);
            false
        } else {
            self.class_list_add(node, class);
            true
        }
    }

    /// `classList.contains(token)`.
    pub fn class_list_contains(&self, node: NodeId, class: &str) -> bool {
        self.class_tokens(node).iter().any(|t| t == class)
    }

    /// `element.style.removeProperty(property)` — returns the value it had.
    ///
    /// Not the same as setting it to `""`: a declaration present with an empty
    /// value still SHADOWS the cascade, while a removed one lets the cascade
    /// through again.
    pub fn remove_style_property(&mut self, node: NodeId, property: &str) -> Option<String> {
        let property = if crate::css::is_custom_property(property) {
            property.trim().to_string()
        } else {
            property.to_ascii_lowercase()
        };
        if node == DOCUMENT {
            return self.document_style.remove(&property);
        }
        self.nodes.get_mut(&node).and_then(|n| n.style.remove(&property))
    }

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
    ///
    /// The **cascade**: user-agent declarations first, the author's layered on
    /// top. That order is the whole rule — `CssProperties::merge` keeps every
    /// field the author did not specify, which is what lets a `<strong>` be
    /// bold and still take `font-weight: normal` from a program.
    /// A CLONE of the computed properties.
    ///
    /// Not public: `getComputedStyle` is the web spelling and it is right next
    /// door. This exists only as an in-crate convenience for callers that want
    /// an owned copy rather than a borrow.
    pub(crate) fn style_properties(&self, node: NodeId) -> crate::css::CssProperties {
        self.get_computed_style(node).clone()
    }

    /// The stored computed style — a READ, never a derivation.
    /// `getComputedStyle(element)` — the resolved property set after the
    /// cascade.
    pub fn get_computed_style(&self, node: NodeId) -> &crate::css::CssProperties {
        if node == DOCUMENT {
            return &self.document_computed;
        }
        self.nodes
            .get(&node)
            .map(|n| &n.computed)
            .unwrap_or(&self.empty_computed)
    }

    /// Run the cascade for one element and store the answer.
    ///
    /// Inheritance is the FLOOR of the cascade, not a layer in it: an inherited
    /// value is what the element has when nothing — not the UA sheet, not the
    /// author — says otherwise. So the parent's computed style goes in first
    /// and every declaration on this element outranks it. `a { color: #0000ee }`
    /// beating a red `<div>` around it is that ordering, and it is what a
    /// browser does.
    ///
    /// The parent is read from its STORED computed style, which is why this
    /// must run top-down: a parent resolved after its child would hand down
    /// last pass's answer. `restyle_subtree` is the only thing that should call
    /// it in bulk, and it walks in that order.
    fn resolve_computed(&mut self, node: NodeId) {
        // **Two passes, because `font-size` is its own basis' exception.**
        //
        // `em` means the element's own computed `font-size` — but on
        // `font-size` itself that would be circular, so there it means the
        // PARENT's. So the cascade runs once with the parent's size to settle
        // `font-size`, then again with the settled size for everything else.
        // `h1 { font-size: 2em; margin: 0.67em 0 }` is exactly this: `2em` is
        // twice the container's text, and the margin is two thirds of the
        // heading's OWN 32px, not of the 16px it inherited.
        //
        // The second pass costs nothing unless the element actually changed
        // font-size, which is the rare case.
        let parent_size = self
            .node(node)
            .and_then(|n| n.parent)
            .and_then(|parent| self.get_computed_style(parent).font_size)
            .unwrap_or(16.0);
        let root_size = self.document_computed.font_size.unwrap_or(16.0);
        let ctx = crate::css::FontContext {
            em: parent_size,
            rem: root_size,
        };
        let mut props = self.cascade(node, ctx);
        let own_size = props.font_size.unwrap_or(parent_size);
        if own_size != parent_size {
            let settled = props.font_size;
            props = self.cascade(
                node,
                crate::css::FontContext {
                    em: own_size,
                    rem: root_size,
                },
            );
            props.font_size = settled;
        }
        if node == DOCUMENT {
            self.document_computed = props;
        } else if let Some(n) = self.nodes.get_mut(&node) {
            n.computed = props;
        }
    }

    /// Re-run the cascade over `root` and everything under it, top-down.
    ///
    /// Parent before child, without exception — the child's inherited values
    /// are read from the parent's freshly stored answer.
    /// One pass of the cascade for one element.
    ///
    /// **The origin order, and it is the whole of the cascade:** the inherited
    /// floor, then the user-agent sheet, then author RULES in specificity
    /// order, then the inline `style=""`. Each layer only overwrites what it
    /// actually declares, so an author rule beats a UA rule and an inline
    /// declaration beats both — which is why `<strong style="font-weight:400">`
    /// is not bold and `p { color: red }` does not defeat `style="color:blue"`.
    fn cascade(&self, node: NodeId, ctx: crate::css::FontContext) -> crate::css::CssProperties {
        let mut props = self
            .node(node)
            .and_then(|n| n.parent)
            .map(|parent| self.get_computed_style(parent).inherited())
            .unwrap_or_default();
        if let Some(n) = self.node(node) {
            // The UA layer is a fixed table with no `var()` in it, so only the
            // author's declarations need resolving.
            props.merge(&n.ua_style.properties_in(ctx));
        }
        // Author rules. Already sorted by specificity then source order, so a
        // forward pass leaves the winner on top with no comparison here.
        // The document root takes author rules even though it has no `DomNode`
        // of its own — one is only created when something sets an attribute on
        // it. Gating on `nodes` meant the page's own box was the one element a
        // stylesheet could never reach.
        if !self.stylesheet.is_empty() && (node == DOCUMENT || self.nodes.contains_key(&node)) {
            let element = ElementRef { doc: self, node };
            for rule in &self.stylesheet {
                if crate::selector::matches(&element, &rule.selector) {
                    props.merge(
                        &rule.declarations.properties_resolved_in(
                            &|value| self.resolve_value(node, value),
                            ctx,
                        ),
                    );
                }
            }
        }
        if let Some(author) = self.style(node) {
            props.merge(
                &author.properties_resolved_in(&|value| self.resolve_value(node, value), ctx),
            );
        }
        props
    }

    /// Re-read every `<style>` in the document into the author stylesheet.
    ///
    /// Rebuilt wholesale rather than patched per element, because a rule is not
    /// owned by the element it came from — deleting one `<style>` can change
    /// what every element in the document computes to. Cheap: stylesheets are
    /// small and this runs when one is written, not per frame.
    fn rebuild_stylesheet(&mut self) {
        let sheets: Vec<NodeId> = self
            .order
            .iter()
            .copied()
            .filter(|id| self.node(*id).map(|n| n.tag == "style").unwrap_or(false))
            .collect();
        let mut rules = Vec::new();
        for sheet in sheets {
            let text = self.text_content(sheet);
            for (selector_text, block) in crate::css::parse_rules(&text) {
                // An unsupported selector drops the WHOLE rule, per CSS error
                // handling — a `:hover` rule that applied unconditionally would
                // be worse than one that never applies.
                let Some(selectors) = crate::selector::parse_selector_list(&selector_text) else {
                    continue;
                };
                for selector in selectors {
                    let mut declarations = Style::new();
                    for declaration in block.split(';') {
                        if let Some((name, value)) = declaration.split_once(':') {
                            declarations.set(name.trim(), value.trim());
                        }
                    }
                    // A selector list is a shorthand for one rule each, and each
                    // carries ITS OWN specificity: `#a, p { }` is not one rule
                    // with two names.
                    rules.push(StyleRule {
                        specificity: selector.specificity(),
                        selector,
                        declarations,
                        order: rules.len(),
                    });
                }
            }
        }
        rules.sort_by_key(|rule| (rule.specificity, rule.order));
        // **Does anything here depend on sibling POSITION?** Asked once, of the
        // sheet, because the answer decides whether inserting a child has to
        // restyle its siblings — the one invalidation a per-node restyle can
        // never cover, since the node that stopped matching is not the node
        // that changed. A document using no `:nth-child()` pays nothing.
        self.positional_rules = rules.iter().any(|rule| rule.selector.is_positional());
        self.stylesheet = rules;
        // Every element in the document may now compute differently.
        self.restyle_subtree(DOCUMENT);
    }

    /// **A structural change moves the SIBLINGS, not the node that changed.**
    ///
    /// Appending a child makes the previous `:last-child` stop matching, and
    /// `:nth-child(odd)` re-colours every row after an insertion. Neither is
    /// reachable by restyling the node that arrived — it is the ones already
    /// there whose match changed — so the whole sibling list is re-cascaded.
    ///
    /// Gated on the sheet actually using a positional selector, because this is
    /// O(siblings) per insertion and would otherwise make building a long list
    /// quadratic for no benefit.
    fn restyle_positional_siblings(&mut self, node: NodeId) {
        if !self.positional_rules {
            return;
        }
        let Some(parent) = self.nodes.get(&node).and_then(|n| n.parent) else {
            return;
        };
        for sibling in self.child_nodes(parent) {
            if sibling != node && self.is_element(sibling) {
                self.restyle_subtree(sibling);
            }
        }
    }

    /// Re-run the cascade over `root` and everything under it, and tell each
    /// widget what actually moved.
    ///
    /// **One walk, and no guessing about what to re-apply.** Two different
    /// declarations reach down the tree — a custom property, which changes what
    /// `color: var(--brand)` three levels down *means*, and an inherited
    /// property, which changes what a descendant that declared nothing at all
    /// *is*. Both were applied long before the write that invalidated them.
    /// Before there was a stored computed style, answering them meant
    /// re-applying a guessed set: every declaration mentioning `var(`, plus
    /// every inherited property, whether or not anything had moved. Now the
    /// cascade re-resolves and the *difference* is the answer.
    ///
    /// Parent before child, without exception — the child's inherited values
    /// are read from the parent's freshly stored answer, so a parent resolved
    /// after its child would hand down last pass's.
    fn restyle_subtree(&mut self, root: NodeId) {
        let mut stack = vec![root];
        while let Some(id) = stack.pop() {
            let before = self.get_computed_style(id).clone();
            self.resolve_computed(id);
            let changes = crate::css::changed_declarations(&before, self.get_computed_style(id));
            // A run carries a RESOLVED style, so a cascade change on an inline
            // element is a change to its parent's line — there is no widget of
            // its own for the declaration to reach.
            if !changes.is_empty() && self.is_inline_content(id) {
                if let Some(parent) = self.node(id).and_then(|n| n.parent) {
                    self.rebuild_inline_content(parent);
                }
            }
            // And a BOX whose own style moved has to re-derive its line too:
            // its text nodes declare nothing, so their runs are built from this
            // element's computed style and nothing else would notice it changed.
            if !changes.is_empty() && self.has_inline_content(id) {
                self.rebuild_inline_content(id);
            }
            if !changes.is_empty() {
                // The cached edges were resolved against the OLD values, and
                // the box arms below read that cache rather than the string
                // passed in — the same order `set_style_property` uses, and for
                // the same reason.
                self.resolve_box_edges(id);
                for (property, value) in changes {
                    self.apply_style_property(id, property, &value);
                }
            }
            stack.extend(self.child_nodes(id));
        }
    }

    /// Is this element inline content of its parent rather than a box?
    ///
    /// **The formatting context comes from the cascade**, which is the whole
    /// reason `display` was put in the UA stylesheet: `control_kind` froze a
    /// Rust type at creation, so a `<span>` was a leaf label for ever and no
    /// declaration could change it.
    ///
    /// Restricted to elements whose widget is a plain text leaf. `<a>` USED to
    /// be excluded, because dissolving it into a run of its parent's text
    /// deleted the click behaviour the link-label widget owned. Runs carry
    /// their source element now (`InlineRun::source`) and the box that paints
    /// them hit-tests them (`FlowLayoutPanel::run_rects`), so the click
    /// survives and the exclusion is gone — a link flows with the sentence
    /// instead of being a box parked on the line.
    fn is_inline_content(&self, node: NodeId) -> bool {
        let Some(n) = self.node(node) else {
            return false;
        };
        let props = self.get_computed_style(node);
        // **Blockification** — CSS 2.1 §9.7. Taking an element out of flow
        // makes its computed `display` block, whatever it was declared as: an
        // absolutely positioned `<span>` is a box with a rect, not a run of its
        // parent's line, because there is no line for it to be part of once it
        // has left the flow. Same for a box that positions itself by
        // `left`/`top` without saying so, which is every VCL control.
        if props.is_out_of_flow() || self.positions_itself(node) {
            return false;
        }
        // The real question is "does this element's widget contribute anything
        // a run would lose?" — which for a plain text leaf is nothing.
        //
        // `<a>` is named alongside it because that answer CHANGED. Its widget
        // owned a click, so dissolving it into a run used to delete one; runs
        // carry their source and are hit-tested now, so the click survives the
        // dissolve and a link can finally flow inside its sentence. It cannot
        // be folded into the `control_kind` test: an `<a>` that BLOCKIFIES must
        // still answer `linklabel` there, because a box gets a widget and a
        // `Label` would emit nothing.
        props.display == Some(crate::css::Display::Inline)
            && (control_kind(&n.tag, &n.input_type) == "label" || n.tag == "a")
    }

    /// Move an element between being a run and being a box, if the cascade
    /// just changed which one it is.
    ///
    /// **Blockification is not only a creation-time question.** `position:
    /// absolute` is usually declared after the element is already in the tree —
    /// it is the ordinary way every VCL frontend writes a form — so an element
    /// that joined as a run has to become a box when it leaves the flow, and
    /// back again if it returns to it. Answering only at `append_child` left a
    /// positioned `<span>` dissolved into its parent's line, painting nothing.
    ///
    /// The current state is read from where the WIDGET is rather than from a
    /// flag beside it: a run's control sits in `detached` while its node has a
    /// parent, and that combination is what a run IS.
    fn reconcile_box_or_run(&mut self, node: NodeId) {
        let Some(parent) = self.node(node).and_then(|n| n.parent) else {
            return;
        };
        let should_be_run = self.is_inline_content(node) && self.holds_inline_content(parent);
        let is_run = self.detached.contains_key(&node);
        if should_be_run == is_run {
            return;
        }
        if should_be_run {
            // Box → run: lift the control out of its container and let the
            // parent's line carry it instead.
            if let Some(widget) = self.extract_widget(node) {
                self.detached.insert(node, widget);
            }
        } else if let Some(widget) = self.detached.remove(&node) {
            // Run → box: give it back a control in the tree. Anything it left
            // behind in the parent's line goes with it.
            if let Some(returned) = self.insert_widget(parent, widget, None) {
                self.detached.insert(node, returned);
                return;
            }
            self.apply_declared_geometry(node);
        }
        self.rebuild_inline_content(parent);
    }

    /// Can this element hold inline content — an inline formatting context?
    ///
    /// Only the box widget renders runs. A `<button>` is inline-level and
    /// text-bearing but is a LEAF control, so `<button><span>x</span></button>`
    /// still reports failure rather than quietly dissolving the span into a
    /// button that cannot draw it. That is a real limitation of the widget set,
    /// not of the model — and reporting it is what stops it being silent.
    fn holds_inline_content(&self, parent: NodeId) -> bool {
        self.node(parent)
            .map(|n| control_kind(&n.tag, &n.input_type) == "flowlayoutpanel")
            .unwrap_or(false)
    }

    /// Re-derive `parent`'s inline runs from its inline children.
    ///
    /// Called wherever one of those children can change — joining the box,
    /// leaving it, its text, or its computed style. One function, called from
    /// each, rather than each site building runs its own way.
    /// Is this child an **atomic inline-level box** — a widget that belongs on
    /// its parent's line rather than on a line of its own?
    ///
    /// The complement of [`Document::is_inline_content`]: that one is "inline
    /// and dissolves into a run of text", this one is "inline-LEVEL and keeps
    /// its own box". `<input>`, `<button>`, `<img>` and `<progress>` are
    /// `inline-block` in the UA sheet — replaced or form elements, which have a
    /// width and a height of their own AND sit in a line of prose. That is
    /// exactly CSS's "atomic inline-level box".
    ///
    /// Out-of-flow is excluded for the same reason it is everywhere else: a
    /// positioned box has left the line, so it is not on it.
    fn is_atomic_inline(&self, node: NodeId) -> bool {
        if !self.is_element(node) || self.is_inline_content(node) {
            return false;
        }
        let props = self.get_computed_style(node);
        if props.is_out_of_flow() || self.positions_itself(node) {
            return false;
        }
        matches!(
            props.display,
            Some(crate::css::Display::InlineBlock)
                | Some(crate::css::Display::InlineFlex)
                | Some(crate::css::Display::InlineGrid)
                // `inline-table` is the table's atomic form and belongs here
                // for the same reason the other three do: inline-level
                // OUTSIDE, its own formatting context INSIDE.
                | Some(crate::css::Display::InlineTable)
        )
    }

    fn rebuild_inline_content(&mut self, parent: NodeId) {
        // **Document order**, and that is the whole of interleaving. Walking
        // the children in the order they appear puts `a <b>B</b> c` back
        // together as three runs; collecting the inline ELEMENTS first and
        // appending the box's own text before them is what made the trailing
        // ` c` inexpressible.
        // Inline-LEVEL widgets are in this list too, as atomic slots. An
        // `<input>` next to a `<label>` shares the label's line in HTML, and it
        // can only do that if it is IN the sequence — collecting text alone and
        // laying the widgets out separately is what put a label on top of its
        // own field.
        let inline: Vec<NodeId> = self
            .child_nodes(parent)
            .into_iter()
            .filter(|c| {
                self.is_text_node(*c) || self.is_inline_content(*c) || self.is_atomic_inline(*c)
            })
            .collect();
        if inline.is_empty() && !self.has_inline_content(parent) {
            return;
        }
        let mut runs = Vec::with_capacity(inline.len());
        for child in inline {
            // **An atomic slot carries no text.** Its size is the widget's, so
            // there is nothing to collapse, transform or shape — the line walk
            // reads the box instead. Emitted in document order with the text
            // runs, which is the whole point: order is what interleaving IS.
            if self.is_atomic_inline(child) {
                let mut font = crate::ide_text::FontSpec::sans(14.0);
                font.apply_computed(self.get_computed_style(child));
                runs.push(crate::layout::InlineRun {
                    text: String::new(),
                    font,
                    color: (0, 0, 0, 255),
                    source: Some(Self::widget_name(child)),
                    atomic: Some(Self::widget_name(child)),
                    cursor: self.get_computed_style(child).cursor,
                });
                continue;
            }
            let raw = if self.is_text_node(child) {
                self.text_data(child)
            } else {
                self.text_content(child)
            };
            // **Whitespace collapsing — CSS Text §white-space: normal.**
            //
            // Source markup is written for humans: newlines and indentation
            // between tags are formatting, not content. Without this,
            // `<p>\n  a <b>B</b>\n</p>` renders with the author's line breaks
            // and indentation in the middle of the sentence, which is the
            // single most visible way a renderer stops looking like HTML.
            //
            // Collapsing is a property of the LINE, not of one run: the space
            // ending `"a "` and the space starting `" c"` are one space
            // between them, so the boundary is carried across runs rather
            // than each being trimmed alone. Leading whitespace at the start
            // of the line and trailing at the end are dropped entirely.
            let text = collapse_whitespace(&raw, runs.is_empty());
            if text.is_empty() {
                continue;
            }
            // A text node has no style of its own — it takes its parent's,
            // which is not a shortcut but the definition: `color` on a `<p>`
            // colours the `<p>`'s text because the text node inherits, and a
            // text node inherits everything since it declares nothing.
            let styled = if self.is_text_node(child) {
                parent
            } else {
                child
            };
            let props = self.get_computed_style(styled);
            let mut font = crate::ide_text::FontSpec::sans(14.0);
            font.apply_computed(props);
            let color = props
                .color
                .map(|c| {
                    let (r, g, b, a) = crate::layout::parse_color(&crate::css::serialize_color(c))
                        .unwrap_or((0, 0, 0, 255));
                    (r, g, b, a)
                })
                .unwrap_or((0, 0, 0, 255));
            // `text-transform` re-cases what is PAINTED, after collapsing and
            // after the style is resolved — never the stored `textContent`, so
            // reading the element back still answers what the author wrote.
            let text = match props.text_transform {
                Some(transform) => transform.apply(&text),
                None => text,
            };
            // **Who this run belongs to.** A text node's run is the BOX's own
            // text and belongs to nobody; an inline element's run carries that
            // element's name, which is what a click on it can be reported
            // against. Without it a run is anonymous paint and an `<a>` inside
            // a sentence has nothing for a hit test to find.
            let source = if self.is_text_node(child) {
                None
            } else {
                Some(Self::widget_name(child))
            };
            runs.push(crate::layout::InlineRun {
                text,
                font,
                color,
                source,
                // Text, not a box — the shaped advance is its width.
                atomic: None,
                cursor: props.cursor,
            });
        }
        self.inline_content.insert(parent, !runs.is_empty());
        self.command(
            parent,
            &WidgetCommand::Custom("SetInlineContent".into(), CommandValue::Runs(runs)),
        );
        self.size_to_content(parent);
    }

    /// **A percentage is not a length until the box has a containing block.**
    ///
    /// `width: 100%` is resolved to pixels when the declaration arrives, and a
    /// frontend that builds its tree bottom-up declares it while the element is
    /// still DETACHED — no parent, so `containing_block` falls back to the
    /// initial one and the percentage freezes against the viewport. A Flutter
    /// button four levels down came out 800x600 and covered the window.
    ///
    /// Re-resolved here, on the way in, because that is where the containing
    /// block is known: the parent's width was settled before the recursion, and
    /// a definite height (declared, or flexed per §9.9) is already on the box.
    /// Against an indefinite parent the percentage resolves to `auto`, which is
    /// CSS's own answer and leaves the content sizing below to decide.
    ///
    /// Only percentages: a length in `px` means the same detached or attached,
    /// and re-applying every declaration on every pass would be churn.
    /// Resolved against the CONTAINING BLOCK, read live — the parent's flow has
    /// just been re-run by the caller, so its rect is current on both axes.
    /// Passing the extents down instead looks right and is not: a stretched
    /// flex item has a perfectly definite used height that no declaration
    /// records, so the caller would offer `0` and the box would keep whatever
    /// stale pixel value it was created with.
    fn reresolve_percentages(&mut self, node: NodeId) {
        for property in ["width", "height"] {
            let is_percent = self
                .style(node)
                .map(|s| s.get(property).trim().ends_with('%'))
                .unwrap_or(false);
            if !is_percent {
                continue;
            }
            // **A percentage height against an indefinite containing block is
            // `auto`** — §10.5. Not "fill the parent's current pixels": the
            // parent has no answer yet, so the box sizes to its own content and
            // the content pass below decides.
            //
            // The inline axis needs no such guard, because a block container's
            // width is always definite.
            if property == "height"
                && self
                    .nodes
                    .get(&node)
                    .and_then(|n| n.parent)
                    .map(|p| p != DOCUMENT && self.axis_mode(p, false) != "declared")
                    .unwrap_or(false)
            {
                continue;
            }
            self.reapply_constrained_axis(node, property);
        }
    }

    /// Settle one box and everything under it — **widths down, heights up**.
    ///
    /// The two halves of sizing run in opposite directions, which is why they
    /// cannot be one walk: a block's width comes from its containing block, so
    /// it is known before its content is looked at, and its height comes from
    /// that content, so it is not known until every descendant has one. Width
    /// on the way in, height on the way out.
    ///
    /// **This used to stop after one level**, and the body's children were the
    /// only boxes that were ever re-measured. A subtree's heights were settled
    /// once each, as it was built, bottom-up — and then every later layout pass
    /// overwrote them with sizes derived from the container, with nothing to
    /// push back. Each ancestor added on top ratcheted the whole subtree
    /// smaller: a Flutter calculator measured 32.5px per row while it was a
    /// bare column, 25.2 once a sibling joined it, 14.0 inside a `Scaffold`,
    /// 12.4 inside a `MaterialApp` — a text line's worth, for a 28px button.
    fn size_subtree(&mut self, node: NodeId, available: f32) {
        // **Whether this box's block size is definite can only be known once
        // its ancestors are.** It is sent at append time, when a bottom-up
        // frontend is still holding a detached subtree — no parent, so nothing
        // above it is definite yet and every box answered `auto`. Re-asked here,
        // where the chain above is settled, so a `Scaffold` that has since been
        // given the window's height tells its column to divide it.
        let mode = self.axis_mode(node, false).to_string();
        self.command(
            node,
            &WidgetCommand::Custom("SetHeightMode".into(), CommandValue::Text(mode)),
        );
        self.reresolve_percentages(node);
        self.fill_available_width(node, available);
        // **Flow it once on the way DOWN too**, now that its own width is
        // settled. A block child's width comes from `fill_available_width`
        // below, but a FLEX or GRID child's comes from its container's
        // layout — and that only runs when the container's rect is set. Doing
        // it only on the way up left every item still carrying its creation
        // width while its own children resolved percentages against it: a
        // Flutter key read `width: 100%` off a cell that had not been narrowed
        // yet and came out a full viewport wide, painting its label into the
        // next column.
        self.reflow(node);
        // The content box the children resolve their own widths against.
        let edges = self.box_edges(node);
        let frame = edges.border.horizontal() + edges.padding.horizontal();
        let inner = (self.border_rect(node).w - frame).max(0.0);
        for child in self.child_nodes(node) {
            if !self.is_element(child) {
                continue;
            }
            // The same boxes the body's own pass skips: a docked or
            // self-positioning child is not in this flow, and one that renders
            // nothing has no box to size.
            let skip = self
                .node(child)
                .map(|n| n.dock.is_some() || renders_nothing(&n.tag, &n.input_type))
                .unwrap_or(true);
            if skip || self.positions_itself(child) {
                continue;
            }
            self.size_subtree(child, inner);
        }
        // **Re-run this box's own flow before reading what it used.**
        // `ContentHeight` is a READ of the last pass — that is what keeps it
        // from disagreeing with where the children actually are — so a box
        // whose children just changed size has to flow them again first, or it
        // answers with the arrangement they had on the way in.
        self.reflow(node);
        self.apply_content_height(node);
    }

    /// Re-run a box's own layout, without changing its size.
    ///
    /// Setting the rect it already has is what asks a panel to arrange its
    /// children again — the same move [`Document::relayout_parent`] makes one
    /// level up.
    fn reflow(&mut self, node: NodeId) {
        if let Some(w) = self.widget_mut(node) {
            let rect = w.rect();
            w.set_rect(rect);
        }
    }

    /// Size a box to its content and re-run the flow it sits in.
    ///
    /// The pair is one operation from a caller's point of view — a box that
    /// grew moved everything after it, and nothing else re-runs the flow — but
    /// [`Document::relayout_body`] sizes its children itself and must not
    /// re-enter, which is why the measuring half is separate.
    fn size_to_content(&mut self, node: NodeId) {
        if self.apply_content_height(node) {
            self.relayout_parent(node);
        }
    }

    /// **A block container is as tall as the boxes it flowed** — §10.6.3 again,
    /// with line boxes made of boxes rather than of text.
    ///
    /// The number comes from the flow itself rather than from a second walk
    /// over the children: the flow is what decided which boxes shared a line
    /// and which took a row, so it is the only thing that knows where the last
    /// one ended. Re-deriving it here would be a second implementation of
    /// wrapping, free to disagree with where the children actually are.
    ///
    /// Only a normal-flow container answers. A flex container sizes its
    /// children to ITSELF — that is what `align-items: stretch` means — so
    /// asking it for a content height gets its own height back, and setting
    /// that as its height would be a fixed point around whatever guess it
    /// started with.
    /// Where this box's flow finished — the measuring half, with no write.
    ///
    /// `None` when the box flowed nothing, for the same reason an empty text
    /// box keeps its guess: collapsing every unstyled container to zero is a
    /// different change from this one.
    fn flowed_height(&mut self, node: NodeId) -> Option<f32> {
        let height = match self.command(
            node,
            &WidgetCommand::Custom("ContentHeight".into(), CommandValue::None),
        ) {
            CommandValue::Number(h) => h as f32,
            _ => return None,
        };
        (height > 0.0).then_some(height)
    }

    fn apply_flowed_height(&mut self, node: NodeId) -> bool {
        let Some(height) = self.flowed_height(node) else {
            return false;
        };
        let Some(w) = self.widget_mut(node) else {
            return false;
        };
        let rect = w.rect();
        if (rect.h - height).abs() < 0.5 {
            return false;
        }
        w.set_rect(LayoutRect::new(rect.x, rect.y, rect.w, height));
        true
    }

    /// **Height from content** — CSS 2.1 §10.6.3.
    ///
    /// A block box with `height: auto` whose children are all inline is as tall
    /// as its line boxes, and how many of those there are depends on where the
    /// text wrapped, which depends on the width it was given. That is the whole
    /// rule, and it is the half of intrinsic sizing that a known width makes
    /// answerable: `width: auto` on a block fills its containing block, so the
    /// width is decided before the content is measured and only the height comes
    /// back from it.
    ///
    /// The other half — shrink-to-fit, where the width comes from the content
    /// too — needs min-content and max-content and is not here. Boxes that need
    /// it (floats, inline-blocks, table cells) keep [`default_size`]'s guess.
    ///
    /// A box with NO inline content keeps its guess too, rather than collapsing
    /// to nothing. CSS says an empty block box is zero high and that is the
    /// right answer, but reaching it means every unstyled container disappears,
    /// which is a change of a different size from this one. Recorded as the
    /// boundary of this pass, not as the rule.
    ///
    /// Answers whether the height CHANGED, so the caller decides what re-runs.
    ///
    /// Measured through the same shaping the renderer paints with
    /// ([`crate::ide_text::measure_rich_text`]), because a height derived from a
    /// line layout nobody draws is a number that happens to be wrong.
    fn apply_content_height(&mut self, node: NodeId) -> bool {
        // A declared height is the author's answer; content does not overrule
        // it. `None` here IS `auto` — the cascade has already folded the UA
        // layer in, so this asks about the used value rather than about one
        // declaration store.
        // **Definite, not merely DECLARED.** A flex or grid item's block size
        // comes from its container's algorithm and no declaration records it,
        // so testing for a declaration let content sizing overrule the share a
        // `Column` had just handed an `Expanded` — every row grew back to its
        // content and the calculator overflowed the window again. `axis_mode`
        // is the one place that knows all three ways a block size becomes
        // definite: declared, flexed, or stretched.
        if self.axis_mode(node, false) == "declared" {
            return false;
        }
        // A box holding element children as well as text is a block container
        // with mixed content, and CSS wraps the text in anonymous block boxes
        // around them. Those do not exist here, so the inline runs are not the
        // whole of this box's content and sizing to them alone would collapse
        // the children out of view.
        if self
            .child_nodes(node)
            .into_iter()
            .any(|c| self.is_element(c) && !self.is_inline_content(c))
        {
            // Unless the children are BOXES this box laid out itself, in which
            // case its content height is exactly where its flow finished — the
            // same rule as §10.6.3, with line boxes made of boxes instead of
            // text. Asking the flow rather than re-deriving it is what keeps
            // the height and the children's positions from disagreeing.
            return self.apply_flowed_height(node);
        }
        // **A line box is at least as tall as the inline-level BOXES on it** —
        // §10.8. Everything left here is inline content, but inline content is
        // not only text: an `inline-block` or a replaced element sits on the
        // line as a box with a height of its own, and measuring the line as if
        // it were text alone answers one text line for a box three times that.
        //
        // That was the whole of the Flutter collapse. A `Padding` wrapping an
        // `ElevatedButton` is `<div><button>…</button></div>`, the button is
        // `inline-block`, so the div measured its caption's ONE LINE — about
        // 17px — for a 28px button. Every ancestor then sized to that: four
        // rows 17px apart, buttons overlapping at half height. The buttons were
        // never mispositioned; the boxes around them were measured as text.
        //
        // Only when there IS such a box. An empty container's flow reports its
        // padding, and taking that as a height would collapse every unstyled
        // `<div>` to nothing — the change this one is deliberately not.
        let has_boxes = self
            .child_nodes(node)
            .into_iter()
            .any(|c| self.is_element(c));
        let boxed = if has_boxes {
            self.flowed_height(node).unwrap_or(0.0)
        } else {
            0.0
        };
        // The runs as the WIDGET holds them — what will actually be painted.
        // Re-deriving them here would let the measurement agree with itself
        // while the renderer drew something else.
        let runs = self.inline_runs(node);
        if runs.is_empty() && boxed <= 0.0 {
            return false;
        }
        let edges = self.box_edges(node);
        let frame_w = edges.border.horizontal() + edges.padding.horizontal();
        let frame_h = edges.border.vertical() + edges.padding.vertical();
        // The widget's rect is the BORDER box on both axes, whatever
        // `box-sizing` says — that only decides how a DECLARED length maps onto
        // it, and this box has no declared height.
        let rect = self.border_rect(node);
        let content_w = (rect.w - frame_w).max(0.0);
        let spans: Vec<_> = runs
            .iter()
            .map(|run| {
                let (r, g, b, a) = run.color;
                (
                    run.text.clone(),
                    run.font.clone(),
                    cosmic_text::Color::rgba(r, g, b, a),
                )
            })
            .collect();
        // Measured the way it will be PAINTED. A box that may not wrap is one
        // line high however narrow it is, and measuring it against the content
        // width would give it the height of text that breaks — the number the
        // renderer then contradicts.
        let wrap = (self.get_computed_style(node).white_space
            != Some(crate::css::WhiteSpace::Nowrap)
            && self.get_computed_style(node).white_space != Some(crate::css::WhiteSpace::Pre))
        .then_some(content_w);
        let (_, content_h) = crate::ide_text::measure_rich_text(&spans, wrap);
        // The taller of the two, not one or the other: a line carrying text
        // AND a box is as tall as whichever reaches further. `boxed` already
        // includes this box's own padding, which is why it is compared against
        // the framed text height rather than added to it.
        let height = (content_h + frame_h).max(boxed);
        if (height - rect.h).abs() < 0.5 {
            return false;
        }
        if let Some(w) = self.widget_mut(node) {
            let r = w.rect();
            w.set_rect(LayoutRect::new(r.x, r.y, r.w, height));
        }
        true
    }

    /// Whether this box was last told it has inline content — so emptying it
    /// still sends the (now empty) list rather than leaving the last one drawn.
    fn has_inline_content(&self, parent: NodeId) -> bool {
        self.inline_content.get(&parent).copied().unwrap_or(false)
    }

    /// A box's inline runs, as the widget currently holds them.
    ///
    /// Read from the WIDGET rather than re-derived, because the question this
    /// answers is "what will be painted", and a re-derivation would agree with
    /// itself while the widget painted something else.
    pub fn inline_runs(&mut self, node: NodeId) -> Vec<crate::layout::InlineRun> {
        match self.command(node, &WidgetCommand::Custom("GetInlineContent".into(), CommandValue::None))
        {
            CommandValue::Runs(runs) => runs,
            _ => Vec::new(),
        }
    }

    /// An element's children, including the document's.
    ///
    /// `DOCUMENT` is the form and has no `DomNode`, so nothing records its
    /// child list — the only truth for a top-level element is its own `parent`.
    /// Every walk from the root has to know that or it stops at the document
    /// and silently covers nothing, which is what a subtree restyle starting at
    /// the body would have done.
    /// `node.childNodes` — every child, in document order, of every kind.
    ///
    /// The LIVE list, answered per call. It is deliberately not a property on
    /// a node handle: a handle is made once and a tree changes, so a stamped
    /// copy would be right when it was taken and wrong immediately after.
    pub fn child_nodes(&self, parent: NodeId) -> Vec<NodeId> {
        // The document's list lives on the document — see `root_children`.
        if parent == DOCUMENT {
            return self.root_children.clone();
        }
        if let Some(n) = self.nodes.get(&parent) {
            return n.children.clone();
        }
        Vec::new()
    }

    /// Put `child` in `parent`'s child list, before `before` or at the end.
    ///
    /// The two ends of the tree keep their lists in different places — a
    /// `DomNode` for an element, [`Self::root_children`] for the document — so
    /// every linking site would otherwise have to remember which. One
    /// function, so none of them does, and so an ordered insert is available
    /// everywhere rather than only where someone wrote it out.
    fn link_child(&mut self, parent: NodeId, child: NodeId, before: Option<NodeId>) {
        let list = if parent == DOCUMENT {
            &mut self.root_children
        } else if let Some(p) = self.nodes.get_mut(&parent) {
            &mut p.children
        } else {
            return;
        };
        // A reference node that is not a child is not a position. Appending is
        // the safe direction and matches what every caller means; the spec's
        // `NotFoundError` is reported by `insert_before` itself, which checks
        // before it gets here.
        match before.and_then(|r| list.iter().position(|c| *c == r)) {
            Some(index) => list.insert(index, child),
            None => list.push(child),
        }
    }

    /// Take `child` out of `parent`'s child list.
    fn unlink_child(&mut self, parent: NodeId, child: NodeId) {
        if parent == DOCUMENT {
            self.root_children.retain(|c| *c != child);
        } else if let Some(p) = self.nodes.get_mut(&parent) {
            p.children.retain(|c| *c != child);
        }
    }

    /// Re-resolve one element's box edges from its declarations.
    ///
    /// Run after **every** style write rather than only after a box property,
    /// because the used value depends on the containing block: a `width` write
    /// on a parent changes what `padding: 10%` means on its child. Resolving
    /// the whole store each time is what keeps declaration order irrelevant,
    /// which is the same reason `min-width` re-runs its axis.
    fn resolve_box_edges(&mut self, node: NodeId) {
        if node == DOCUMENT {
            return;
        }
        let props = self.style_properties(node);
        // Percentages resolve against the containing block's WIDTH on both
        // axes — see `BoxEdges::resolve`.
        let basis = self.containing_block(node).w;
        if let Some(n) = self.nodes.get_mut(&node) {
            n.box_edges = BoxEdges::resolve(&props, basis);
        }
    }

    /// The **border box** — the rectangle the widget occupies, backgrounds
    /// paint, and every existing rect in the tree already means.
    /// The border box, unwrapped.
    ///
    /// NOT public: the web spelling is `getBoundingClientRect()`, and having
    /// both meant two names for one measurement. Kept internally only as the
    /// shorthand the other three boxes are derived from.
    fn border_rect(&mut self, node: NodeId) -> LayoutRect {
        self.get_bounding_client_rect(node).unwrap_or_default()
    }

    /// The **padding box** — inside the border, padding included.
    ///
    /// This is the containing block for absolutely positioned descendants
    /// (CSS 2.1 §10.1), which is the one place the distinction from the border
    /// box is observable today.
    pub fn padding_rect(&mut self, node: NodeId) -> LayoutRect {
        let edges = self.box_edges(node).border;
        inset(self.border_rect(node), edges)
    }

    /// The **content box** — inside the padding. Where a container arranges
    /// its children and where a control draws its own contents.
    pub fn content_rect(&mut self, node: NodeId) -> LayoutRect {
        let edges = self.box_edges(node);
        inset(inset(self.border_rect(node), edges.border), edges.padding)
    }

    /// The **margin box** — the space the element claims from its siblings.
    pub fn margin_rect(&mut self, node: NodeId) -> LayoutRect {
        let edges = self.box_edges(node).margin;
        let r = self.border_rect(node);
        LayoutRect::new(
            r.x - edges.left,
            r.y - edges.top,
            r.w + edges.horizontal(),
            r.h + edges.vertical(),
        )
    }

    /// Does a declared `width` on this element measure its CONTENT?
    ///
    /// CSS's initial value says yes, and that is the default here. The toolkit
    /// convention — a VCL or WinForms `Width` includes border and padding —
    /// is not baked in as an inverted default; it is declared as
    /// `box-sizing: border-box` on the elements it is true of, by the UA
    /// control rule in [`crate::ua`]. That keeps the box model CSS-shaped and
    /// leaves every frontend, Flutter included, free to declare its own.
    fn uses_content_box(&self, node: NodeId) -> bool {
        self.style_properties(node).box_sizing != Some(crate::css::BoxSizing::BorderBox)
    }

    /// The resolved edges. Zero for anything with no box declarations, which
    /// is what makes all four rects coincide — the model this replaced.
    pub fn box_edges(&self, node: NodeId) -> BoxEdges {
        self.nodes
            .get(&node)
            .map(|n| n.box_edges)
            .unwrap_or_default()
    }

    /// `element.style.getPropertyValue(property)`.
    ///
    /// Geometry answers from the **control**, not from the declaration: a rect
    /// is computed, and a docked or laid-out element is not where its own
    /// `left` said it would be. Everything else answers from the declaration
    /// store, which is what makes a property the write side accepts readable —
    /// `color`, `background-color` and `font-size` all returned `""` before,
    /// having been consumed into widget commands and forgotten.
    /// `element.style.getPropertyValue(property)` — the DECLARED value.
    ///
    /// ⛔ This used to answer `left`/`top`/`width`/`height` out of the laid-out
    /// rect, so `setProperty("top", "1em")` read back `"16px"`. That is
    /// `getComputedStyle`'s job, not this one: CSSOM §6.4.2 says a declaration
    /// block serializes what was DECLARED, and nothing reachable from
    /// `element.style` is resolved against layout. The resolved answer moved to
    /// `computed_style_property` below rather than being deleted.
    ///
    /// The split is also what makes the engines interchangeable — htmlbox
    /// always answered the declared value, so a single op could only ever have
    /// been right for one of the two engines.
    pub fn get_style_property(&mut self, node: NodeId, property: &str) -> String {
        let property = property.to_ascii_lowercase();
        self.declared_style(node, &property)
    }

    /// `getComputedStyle(element).getPropertyValue(property)` — the RESOLVED
    /// value, in used units.
    ///
    /// Geometry comes from the laid-out rect, which is the whole point: a
    /// toolkit's `Control.Left` wants the pixel the control actually occupies,
    /// not whatever string was authored. Everything else falls back to the
    /// declared value — an honest floor, since a full computed-value pass would
    /// have to resolve every property against the cascade.
    pub fn computed_style_property(&mut self, node: NodeId, property: &str) -> String {
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
        // ⛔ `select.value = v` selects the option WHOSE VALUE IS `v`
        // (HTML §4.10.7). It is not an index — `selectedIndex` is the separate
        // IDL member for that, and `gui.rs` already routes `SelectedIndex`
        // there and says so. This arm used to `parse()` the string as an index,
        // which made `sel.value = "1"` mean "select the second option" instead
        // of "select the option worth 1", and made `value` a number in a member
        // the IDL declares a DOMString.
        //
        // An item's value here is its TEXT: these lists hold strings, and
        // `option.value` falls back to the option's label when no `value`
        // attribute is present — which is always, for a widget-backed list.
        if matches!(self.value_kind(node), ValueKind::Index) {
            let mut index = 0usize;
            while let CommandValue::Text(item) = self.command(node, &WidgetCommand::GetItem(index))
            {
                if item == value {
                    self.command(node, &WidgetCommand::SetSelectedIndex(index));
                    return;
                }
                index += 1;
            }
            // No option carries that value. HTML deselects, but the command
            // vocabulary has no "select nothing" — `SetSelectedIndex` takes a
            // `usize` — so the selection is left as it was. Stated rather than
            // silently approximated with index 0, which would SELECT something
            // the assignment never named.
            return;
        }

        let cmd = match self.value_kind(node) {
            ValueKind::Text => WidgetCommand::SetText(value.to_string()),
            ValueKind::Number => WidgetCommand::SetValue(value.trim().parse().unwrap_or(0.0)),
            // Handled above; `value_kind` is re-matched here only for
            // exhaustiveness.
            ValueKind::Index => return,
            ValueKind::Checked => {
                WidgetCommand::SetChecked(!value.eq_ignore_ascii_case("false") && !value.is_empty())
            }
        };
        self.command(node, &cmd);
    }

    pub fn value(&mut self, node: NodeId) -> String {
        // The getter half of the note on `set_value`: a select answers with
        // its SELECTED OPTION'S VALUE, not its index. `GetValue` on a list
        // control reports `CommandValue::Index`, which the arm below would
        // stringify — so `sel.value` used to read back `"1"` where HTML says
        // the option's own text.
        if matches!(self.value_kind(node), ValueKind::Index) {
            let selected = self.selected_index(node);
            return match usize::try_from(selected) {
                Ok(index) => self.item_text(node, index),
                // `selectedIndex` is -1 when nothing is selected, and `value`
                // is the empty string then.
                Err(_) => String::new(),
            };
        }
        match self.command(node, &WidgetCommand::GetValue) {
            CommandValue::Text(s) => s,
            CommandValue::Number(n) => format_number(n),
            CommandValue::Index(i) => i.to_string(),
            // A checkbox's `value` is its SUBMISSION value, not its state —
            // `checked` is the state — and HTML's default for it is `"on"`
            // whether or not the box is ticked.
            CommandValue::Bool(_) => "on".to_string(),
            CommandValue::Color(r, g, b, _) => format!("#{:02x}{:02x}{:02x}", r, g, b),
            // Inline content is not a submission value — no control answers
            // `GetValue` with runs, and a box that has them is not a control.
            CommandValue::Runs(_) => String::new(),
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
        // **The spec's own definition**: setting `textContent` replaces all
        // children with a single text node. For a box that is exactly right and
        // is what makes `p.textContent = "x"` and building the same tree by
        // hand agree.
        //
        // A BORDERED box is the exception, and HTML makes the same one: its
        // text is a `<legend>`, drawn in the gap the border leaves across the
        // top edge rather than as content of the line. It keeps the command
        // path, which is also what `TGroupBox.Caption` writes to.
        //
        // Leaf controls keep it too. A `<button>`'s content model is phrasing
        // in HTML, but the widget here is a leaf that draws one string — a
        // stated limitation of the widget set, not a claim about the tree.
        if self.holds_inline_content(node) && !self.is_bordered(node) {
            let children = self.child_nodes(node);
            for child in children {
                self.remove_child(node, child);
            }
            if !text.is_empty() {
                let data = self.create_text_node(text);
                self.append_child(node, data);
            } else {
                self.rebuild_inline_content(node);
            }
            return;
        }
        self.apply_text(node, text);
    }

    /// Give a control the text it draws.
    ///
    /// The half of `textContent` that is not about the tree: the box case above
    /// builds nodes, and everything else — a leaf's caption, a list's items —
    /// ends here. One function, because a text node appended to a leaf has to
    /// reach the control exactly as an assignment to `textContent` does; two
    /// would be two ways for the same content to arrive, differing.
    fn apply_text(&mut self, node: NodeId, text: &str) {
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
        // A `<style>`'s text IS the author stylesheet. Writing it is how a rule
        // enters the cascade, and it can change what every element in the
        // document computes to — not just this one.
        if self.node(node).map(|n| n.tag == "style").unwrap_or(false) {
            self.rebuild_stylesheet();
            return;
        }
        // If this element IS a run of its parent's text, the parent's line just
        // changed. The run holds a copy of the text by necessity — shaping
        // needs it all at once — so the copy is re-derived here rather than
        // left to drift.
        if let Some(parent) = self.node(node).and_then(|n| n.parent) {
            if self.is_inline_content(node) {
                self.rebuild_inline_content(parent);
            }
        }
        // **A table's columns are measured from its cells' content**, so
        // writing text into a cell resizes columns that other rows share. This
        // is the difference between a table and a grid: nobody declared these
        // widths, so nothing else can be relied on to recompute them.
        self.relayout_enclosing_table(node);
    }

    pub fn text_content(&mut self, node: NodeId) -> String {
        if node == DOCUMENT {
            return self.title();
        }
        // Every non-element node answers its own data — `CharacterData.data`
        // for text and CDATA, and the same for a comment and a processing
        // instruction even though those two are excluded from an ANCESTOR's
        // `textContent` below. Both halves are the spec's, and they disagree on
        // purpose.
        if !self.is_element(node) {
            return self.node(node).map(|n| n.data.clone()).unwrap_or_default();
        }
        // **The concatenation of every descendant text node, in tree order** —
        // the spec's definition, and the read that pairs with the write above.
        // A box whose content is `a <b>B</b> c` answers `"a B c"`, which no
        // single-string field could.
        if self.holds_inline_content(node) && !self.is_bordered(node) {
            // Elements recurse; character data contributes; comments and
            // processing instructions do not. `is_character_data` is the
            // spec's line, and CDATA falls on the text side of it because
            // `CDATASection` IS a `Text` (DOM §4.10).
            let children: Vec<NodeId> = self
                .child_nodes(node)
                .into_iter()
                .filter(|child| self.is_element(*child) || self.is_character_data(*child))
                .collect();
            if !children.is_empty() {
                return children
                    .into_iter()
                    .map(|child| self.text_content(child))
                    .collect();
            }
        }
        // A LEAF's text children are its content too — see `append_child`. It
        // has no line to lay them out on, so they are its caption; but they are
        // still the nodes, and the nodes are what `textContent` answers. Asking
        // the widget instead would answer with whatever it was last told, which
        // is the copy rather than the fact.
        // Elements count here too, and recursing is what `textContent` means:
        // a `<button><span>7</span></button>` answers `"7"`. Filtering to
        // character data alone was right when a leaf could only hold text
        // nodes, and wrong the moment it could hold an inline element — which
        // is how Flutter spells every label.
        let runs: Vec<NodeId> = self
            .child_nodes(node)
            .into_iter()
            .filter(|child| self.is_element(*child) || self.is_character_data(*child))
            .collect();
        let data: String = runs
            .into_iter()
            .map(|child| self.text_content(child))
            .collect();
        if !data.is_empty() {
            return data;
        }
        match self.command(node, &WidgetCommand::GetText) {
            CommandValue::Text(s) => s,
            _ => String::new(),
        }
    }

    /// Does this box draw its text as a `<legend>` rather than as content?
    fn is_bordered(&self, node: NodeId) -> bool {
        self.node(node).map(|n| n.tag == "fieldset").unwrap_or(false)
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
    pub fn get_bounding_client_rect(&mut self, node: NodeId) -> Option<LayoutRect> {
        // The DOCUMENT's box is the viewport. `border_rect` handled this and
        // this did not, so the two disagreed for exactly the one node a page
        // is most likely to measure.
        if node == DOCUMENT {
            return Some(self.form.rect());
        }
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
            let option = self.create_element_typed("input", "radio");
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
    /// **Pull the live pseudo-class state and re-cascade what moved.**
    ///
    /// The pointer moves without the document being told: the arranging panel
    /// maintains `hovered()` on its children directly. So the state is pulled
    /// here, once, rather than pushed — and an element whose state CHANGED has
    /// its cascade re-run, because computed values are cached and a `:hover`
    /// rule that only matched on the next unrelated restyle would apply at
    /// random.
    ///
    /// Called from [`Document::drain_events`], which the host already pumps
    /// every frame, so `:hover` follows the pointer without a second clock.
    ///
    /// Only elements that HAVE a widget are asked — a run has no state of its
    /// own, and asking for one would allocate a map entry per text node.
    pub fn refresh_interaction_states(&mut self) -> bool {
        let mut changed = Vec::new();
        for node in self.order.clone() {
            if !self.is_element(node) {
                continue;
            }
            let Some(widget) = self.widget_mut(node) else {
                self.states.remove(&node);
                continue;
            };
            let hover = widget.hovered();
            let checked = matches!(
                self.command(node, &WidgetCommand::GetValue),
                CommandValue::Bool(true)
            );
            let now = (hover, checked);
            if self.states.get(&node) != Some(&now) {
                self.states.insert(node, now);
                changed.push(node);
            }
        }
        for node in &changed {
            self.restyle_subtree(*node);
        }
        !changed.is_empty()
    }

    /// `EventTarget.addEventListener(type, callback)` — DOM §2.7.
    ///
    /// The browser keeps the listener. A host registers by wrapping its own
    /// callback in a closure; nothing here knows what is inside it.
    pub fn add_event_listener(
        &self,
        node: NodeId,
        kind: &str,
        handler: EventHandler,
    ) -> ListenerId {
        self.listeners.add(node, kind, handler)
    }

    /// `EventTarget.removeEventListener` — by the handle `add_event_listener`
    /// issued.
    ///
    /// A HANDLE rather than the callback, because the browser cannot compare
    /// two opaque closures and the spec's rule is identity, not equality. A
    /// host that needs "remove the listener I registered with THIS callback"
    /// keeps its own callback→handle map; that is a host concern, since only
    /// the host can tell its callbacks apart.
    pub fn remove_event_listener(&self, id: ListenerId) {
        self.listeners.remove(id);
    }

    /// Every listener on a node, dropped with the node.
    pub fn remove_all_listeners(&self, node: NodeId) {
        self.listeners.remove_all(node);
    }

    /// `EventTarget.dispatchEvent(event)` — runs the listeners synchronously
    /// and answers `false` if one called `preventDefault`.
    ///
    /// The registry is cloned out first: it is an `Arc`, so this is a refcount
    /// bump, and it is what lets the handlers be borrowed from the table while
    /// the document itself is handed to them mutably.
    pub fn dispatch_event(&mut self, event: &mut DomEvent) -> bool {
        let listeners = self.listeners.clone();
        listeners.dispatch(self, event)
    }

    /// Turn the interactions that happened into DOM events and DISPATCH them.
    ///
    /// The replacement for handing a caller a queue. The browser translates its
    /// own input into `click`/`input`/`change`, then calls the listeners — the
    /// order a browser does it in, and the only order in which `preventDefault`
    /// can still cancel anything.
    ///
    /// Answers how many events were dispatched, so a frame loop can tell
    /// whether anything happened without being handed the events themselves.
    pub fn dispatch_pending_events(&mut self) -> usize {
        let events = self.drain_events();
        let count = events.len();
        for mut event in events {
            self.dispatch_event(&mut event);
        }
        count
    }

    /// The interactions that happened, as DOM events.
    ///
    /// ⚠ NOT a web API — no browser lets a page pull from its event queue; the
    /// UA dispatches. Kept `pub` only until the last caller moves to
    /// `dispatch_pending_events`, and it is the translation half of it: widget
    /// events in, `click`/`input`/`change` out.
    pub fn drain_events(&mut self) -> Vec<DomEvent> {
        self.refresh_interaction_states();
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
                out.push(DomEvent::new(node, kind));
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
    open_document(Document::new(title))
}

/// Open an XML document — names keep their case. See [`DocumentKind`].
pub fn new_xml_document(title: &str) -> DocumentId {
    open_document(Document::new_xml(title))
}

fn open_document(document: Document) -> DocumentId {
    let mut docs = documents().lock().unwrap();
    docs.next_id += 1;
    let id = docs.next_id;
    docs.docs.insert(id, Mutex::new(document));
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
            // The picker was WRITTEN and unreachable: `ColorPicker` implements
            // `PanelWidget`, emits `ColorChanged`, and `dom.rs` already turned
            // that into `input`/`change` — nothing named it, so
            // `<input type="color">` fell to the text arm and rendered a box
            // you could type in. `value_kind_of` already answers `Text` here,
            // so the `#rrggbb` value reaches it as `SetText` unchanged.
            "color" => "colorpicker",
            // A button plus the chosen file's name, opening the platform
            // chooser — which is what a browser draws and what VB's
            // `OpenFileDialog` is. It fell to the text arm before, rendering an
            // editable field that neither picks a file nor reports one.
            "file" => "fileinput",
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
        // **`<a>` answers here only when it is a BOX.** An `<a>` in flow is
        // inline content — a run of its parent's sentence — and never reaches
        // this function; `is_inline_content` names it explicitly.
        //
        // What DOES reach here is a blockified one: `position: absolute`, or
        // the `Left`/`Top` every VCL and WinForms designer writes. That is a
        // `LinkLabel`, and it must stay one — a `Label` has no `handle_mouse`
        // and emits nothing, so answering `label` here would silently delete
        // the click and the pointer cursor from every designer-placed link.
        "a" => "linklabel",
        "img" => "picturebox",
        "canvas" => "canvas",
        // A block container. `Panel`/`GroupBox` are backgrounds that hold no
        // children, so a nestable container is the only honest mapping — and
        // vertical flow IS what CSS `display:block` does to children.
        "fieldset" => "flowlayoutpanel",
        // **A table is a BOX that lays its children out as a table** — the
        // formatting context comes from `display: table` in the UA sheet, not
        // from the widget type. This used to answer `datagridview`, which made
        // an HTML *layout* impersonate a .NET *control*: `<table>` got a
        // DataGrid with columns and a header of its own, and the `<tr>`s and
        // `<td>`s inside it became panels beside a grid that had never heard of
        // them. `DataGridView` is still reachable by the name that means it.
        "table" => "flowlayoutpanel",
        "ul" | "ol" => "listbox",
        "menu" => "menustrip",
        // A `<select>`'s options are the CONTROL's content, not boxes beside
        // it — a browser does not give them a box either, so they stay leaves.
        "option" | "optgroup" => "label",
        // **A cell is a BOX.** These were labels on the reasoning that they are
        // "a container's CONTENT rather than boxes in their own right". That is
        // the WinForms model, not HTML's: a `<td>` may hold a `<div>`, a
        // `<form>` or another table, and `<summary>` takes phrasing and heading
        // content. Mapped to a leaf, `add_child` was REFUSED and every such
        // child went to `detached` — the same silent-label trap that once
        // swallowed `<html><body>` whole.
        //
        // A cell is now a cell IN A ROW as well as a box: `ua.rs` declares
        // `display: table-cell` and `Formatting::Table` reads it. The widget
        // type is unchanged — what a `<td>` needed all along was a container,
        // and what makes it a table cell is the cascade.
        "td" | "th" | "caption" | "legend" | "summary" => "flowlayoutpanel",
        // Table structure: rows and sections hold cells, so they are
        // containers rather than text.
        "tr" | "thead" | "tbody" | "tfoot" | "colgroup" | "col" => "flowlayoutpanel",
        "iframe" | "embed" | "object" => "picturebox",
        // A rule is a thematic break: a thin panel, which is what it draws as.
        // Without an arm here it would fall to `_ => "label"` at 120x20 — the
        // silent-label trap — which is why a separator stayed on `<menu>` in
        // the dotnet adapter until this existed.
        "hr" => "panel",
        // **A block box has text AND children**, which is why these are not
        // labels. `<p>`, `<h1>` and the rest are text-bearing *and* may contain
        // phrasing content: mapped to a leaf, the leaf refused `add_child` and
        // `append_child` put the child back in `detached`, so
        // `<p>a <strong>b</strong></p>` silently lost the `<strong>`.
        //
        // The container draws its own text (`FlowLayoutPanel::caption`, which
        // is wxhtmledit's `Box::ownText`) and arranges its children under it.
        //
        // Still labels, deliberately: `option`, `optgroup` and `label` itself.
        // The first two are a CONTROL's content and get no box in a browser
        // either; `label` is what every VCL `TLabel` maps to, so changing it
        // would touch every form in the corpus. `td`/`th`/`caption`/`legend`/
        // `summary` USED to be here and are boxes now — see the arm above.
        "p" | "h1" | "h2" | "h3" | "h4" | "h5" | "h6" | "blockquote" | "figure" | "figcaption"
        | "address" | "pre" | "center" | "dl" | "dd" | "dt" | "details" => "flowlayoutpanel",
        "dialog" | "div" | "form" | "section" | "article" | "main" | "aside" | "header"
        | "footer" | "nav" | "li" => "flowlayoutpanel",
        // **The document's own three elements.** Nothing created one until a
        // parser did: `document.body` IS the root node here, so a program that
        // builds a form never names them. Parsed markup always does, and
        // without an arm all three fell to `_ => "label"` — a leaf, which
        // refuses `add_child`, so `<html><body>…` dropped the entire page and
        // `<head>` dropped the `<style>` that was about to become the author
        // stylesheet. The silent-label trap swallowing a whole document.
        //
        // `head` is a container that renders nothing — the two are separate
        // facts. `is_hidden_element` already says it draws no box; holding its
        // children is what makes the metadata reachable at all.
        "html" | "head" | "body" => "flowlayoutpanel",
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
        "colorpicker" => "colorpicker",
        "fileinput" => "fileinput",
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
    // The HTML half is the UA stylesheet's hidden-elements rule, and lives
    // there — one list, citing the spec section it came from, rather than a
    // second copy here that could drift from it.
    crate::ua::is_hidden_element(tag)
        // A hidden input is hidden by its TYPE, not its tag, so it is not a
        // tag-selector rule and stays here.
        || (tag == "input" && input_type == "hidden")
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
    // A text-bearing block box is the same widget as a layout panel and nothing
    // like the same shape either: a paragraph is a line of text, not a 200x150
    // region. The tag decides for the same reason `hr` does.
    //
    // This is a starting geometry, not a measurement — the honest answer is a
    // height the box takes from its content, which is intrinsic sizing and does
    // not exist here yet. Recorded so the placeholder is visible rather than
    // looking like a considered default.
    if matches!(
        tag,
        "p" | "h1"
            | "h2"
            | "h3"
            | "h4"
            | "h5"
            | "h6"
            | "blockquote"
            | "figcaption"
            | "address"
            | "pre"
            | "center"
            | "dl"
            | "dd"
            | "dt"
    ) {
        return (200.0, 20.0);
    }
    match kind {
        "textbox" | "combobox" | "numericupdown" | "datetimepicker" => (160.0, 24.0),
        // A colour swatch is a BUTTON, not a field — it shows the colour and
        // opens a picker, so it takes a control's height and none of a text
        // box's width. Chrome's own is about this size.
        "colorpicker" => (64.0, 24.0),
        // Button plus a filename — wider than a field, one line tall.
        "fileinput" => (240.0, 24.0),
        // Tall, because it holds lines — the same call `RICHTEXTBOX_DEF` makes
        // with its own (100, 96). One line high made an unstyled memo look
        // like a text field.
        "richtextbox" => (160.0, 96.0),
        "button" => (96.0, 28.0),
        "checkbox" | "radiobutton" => (140.0, 20.0),
        "trackbar" | "progressbar" => (160.0, 20.0),
        "listbox" | "listview" | "datagridview" | "treeview" => (200.0, 120.0),
        // HTML §4.12.5: a `<canvas>` with no `width`/`height` attribute is
        // 300x150. Not a toolkit preference — a page that draws without sizing
        // the canvas first depends on it, and it was sharing the panel's
        // 200x150 for no reason but adjacency in this list.
        "canvas" => (300.0, 150.0),
        "panel" | "groupbox" | "flowlayoutpanel" | "picturebox" => (200.0, 150.0),
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

// The second length parser used to live here. It answered the same question as
// `css::parse_length` and answered it differently: `em` and `rem` were both a
// hardcoded 16px, and `pt` was 4/3 of the value where CSS says 96/72 — so a
// declaration's used value depended on which of two functions it happened to
// reach. `Document::px` resolves through the one parser, in the element's own
// font context.

/// Collapse a run's whitespace the way `white-space: normal` says.
///
/// Every tab, newline and run of spaces becomes a single space; whitespace at
/// the very start of the line is dropped. `at_line_start` is what makes this
/// composable across runs — the caller passes `true` only for the first run, so
/// `"a "` followed by `" c"` yields one space between them rather than two, and
/// the indentation before the first word disappears while the space inside the
/// sentence survives.
///
/// Trailing whitespace on the LAST run is left alone here: whether a line has a
/// trailing space is only knowable once the line is complete, and it is
/// invisible at the end of a left-aligned line anyway. Stated rather than
/// silently half-done.
fn collapse_whitespace(text: &str, at_line_start: bool) -> String {
    let mut out = String::with_capacity(text.len());
    let mut pending_space = false;
    for ch in text.chars() {
        if ch.is_whitespace() {
            pending_space = true;
            continue;
        }
        if pending_space && !(out.is_empty() && at_line_start) {
            out.push(' ');
        }
        pending_space = false;
        out.push(ch);
    }
    // A run that is nothing but whitespace still separates its neighbours —
    // the newline and indent between `</b>` and the next word are one space,
    // not nothing. Dropping it entirely is how `a<b>B</b>c` appears where the
    // author wrote three separate lines.
    if pending_space && !out.is_empty() {
        out.push(' ');
    } else if pending_space && !at_line_start {
        out.push(' ');
    }
    out
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

/// Shrink a rect by one set of edges — the step from any box to the next one
/// inside it. A box narrower than its own edges collapses to zero rather than
/// inverting, which is what CSS means by a used width flooring at 0.
fn inset(r: LayoutRect, e: Edges) -> LayoutRect {
    LayoutRect::new(
        r.x + e.left,
        r.y + e.top,
        (r.w - e.horizontal()).max(0.0),
        (r.h - e.vertical()).max(0.0),
    )
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
    fn text_transform_recases_the_painted_run_but_not_text_content() {
        // The invariant that makes this a RENDERING property: the run the
        // widget draws is uppercased, and reading the element back still
        // answers exactly what the author wrote. A transform applied to the
        // stored text instead would pass the first assertion and fail this.
        let mut doc = Document::new("t");
        let p = doc.create_element_typed("p", "");
        doc.append_child(DOCUMENT, p);
        doc.set_text_content(p, "hello world");
        doc.set_style_property(p, "text-transform", "uppercase");

        assert_eq!(doc.inline_runs(p)[0].text, "HELLO WORLD");
        assert_eq!(doc.text_content(p), "hello world");
    }

    #[test]
    fn letter_spacing_cascades_onto_the_run_font() {
        // The other half of the path: `ide_text` proves the spacing reaches
        // cosmic-text, this proves the declaration reaches the run's spec —
        // inherited from the container, so neither end is assumed.
        let mut doc = Document::new("t");
        let box_ = doc.create_element_typed("div", "");
        doc.append_child(DOCUMENT, box_);
        let p = doc.create_element_typed("p", "");
        doc.append_child(box_, p);
        doc.set_text_content(p, "spaced");
        doc.set_style_property(box_, "letter-spacing", "4px");

        assert_eq!(doc.inline_runs(p)[0].font.letter_spacing, 4.0);
    }

    #[test]
    fn text_transform_inherits_and_capitalize_titlecases_each_word() {
        // Declared on the CONTAINER, applied to a descendant's run — which is
        // what makes it an inherited property rather than one the element must
        // declare for itself.
        let mut doc = Document::new("t");
        let box_ = doc.create_element_typed("div", "");
        doc.append_child(DOCUMENT, box_);
        let p = doc.create_element_typed("p", "");
        doc.append_child(box_, p);
        doc.set_text_content(p, "hello wide world");
        doc.set_style_property(box_, "text-transform", "capitalize");

        assert_eq!(doc.inline_runs(p)[0].text, "Hello Wide World");
    }

    #[test]
    fn z_index_reorders_positioned_siblings_and_ignores_static_ones() {
        // Two overlapping boxes. Document order decides while both carry the
        // same z-index, and `z-index` overrules it — but ONLY for positioned
        // boxes, which is why the static pair below does not respond.
        // `<span>` rather than `<div>`: a div is the flow-panel widget, which
        // paints nothing unless it is a `<fieldset>`, so two overlapping divs
        // would prove only that neither of them drew.
        let overlapping = |z: Option<&str>, positioned: bool| {
            let mut doc = Document::new("t");
            let panel = container_with(&mut doc, 100.0, 100.0);
            let under = doc.create_element_typed("span", "");
            doc.append_child(panel, under);
            let over = doc.create_element_typed("span", "");
            doc.append_child(panel, over);
            for (node, colour) in [(under, "#ff0000"), (over, "#0000ff")] {
                if positioned {
                    doc.set_style_property(node, "position", "absolute");
                }
                doc.set_style_property(node, "left", "0px");
                doc.set_style_property(node, "top", "0px");
                doc.set_style_property(node, "width", "60px");
                doc.set_style_property(node, "height", "60px");
                doc.set_style_property(node, "background-color", colour);
            }
            // Lift the FIRST one, so a change of order is unambiguous.
            if let Some(z) = z {
                doc.set_style_property(under, "z-index", z);
            }
            let painted = painted(&mut doc, 100, 100);
            let px = painted.pixel(20, 20).unwrap();
            (px.red(), px.blue())
        };

        let (r, b) = overlapping(None, true);
        assert!(b > r, "document order paints the later sibling on top");

        let (r, b) = overlapping(Some("5"), true);
        assert!(r > b, "a higher z-index lifts the earlier sibling above it");

        // In CSS, `z-index` is inert on a static box — it takes no part in the
        // positioned pass at all. Here it is NOT, and the reason is not
        // `z-index`: a box carrying `left`/`top` is INFERRED to be positioned
        // even when nothing declared `position`, so this box was never static.
        //
        // Recorded as the behaviour it has rather than the behaviour it should
        // have, so removing the inference flips this and says so.
        let (r, b) = overlapping(Some("5"), false);
        assert!(
            r > b,
            "KNOWN DIVERGENCE: coordinates alone make a box positioned here, \
             so z-index reaches it; CSS would leave it static and inert"
        );
    }

    #[test]
    fn a_modal_paints_over_the_page_from_the_top_layer() {
        // The whole point of the top layer: the modal's pass runs AFTER the
        // document's, so its backdrop dims a button that was painted first —
        // and would dim it just the same if the button came later in the
        // tree or carried a large z-index, neither of which the top layer
        // consults.
        let mut doc = Document::new("t");
        let button = doc.create_element_typed("button", "");
        doc.append_child(DOCUMENT, button);
        let dialog = doc.create_element_typed("dialog", "");
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
        // The document is never empty: tree construction has already put
        // `<html>`, `<head>` and `<body>` in it, so the question is whether
        // creating adds to that — not whether the count is zero.
        let before = doc.form().control_count();
        let cb = doc.create_element_typed("input", "checkbox");
        assert!(!doc.is_connected(cb), "createElement must not insert");
        assert_eq!(doc.form().control_count(), before);
        doc.append_child(DOCUMENT, cb);
        assert!(doc.is_connected(cb));
        assert_eq!(doc.form().control_count(), before);
    }

    #[test]
    fn checked_is_the_widgets_state_not_a_copy() {
        let mut doc = Document::new("t");
        let cb = doc.create_element_typed("input", "checkbox");
        doc.append_child(DOCUMENT, cb);
        doc.set_checked(cb, true);
        assert!(doc.checked(cb));
        // Read it straight off the control to prove there is no second copy.
        // Named from the node rather than spelled `"n1"`: the document already
        // holds `<html>`, `<head>` and `<body>`, so the first element a test
        // creates has never been node 1 since tree construction landed.
        let name = Document::widget_name(cb);
        let v = doc.form_mut().send_command(&name, &WidgetCommand::GetValue);
        assert!(matches!(v, CommandValue::Bool(true)));
    }

    #[test]
    fn value_survives_being_set_before_insertion() {
        let mut doc = Document::new("t");
        let input = doc.create_element_typed("input", "text");
        doc.set_value(input, "hello");
        assert_eq!(doc.value(input), "hello");
        doc.append_child(DOCUMENT, input);
        assert_eq!(doc.value(input), "hello");
    }

    #[test]
    fn id_is_an_attribute_not_a_rename() {
        let mut doc = Document::new("t");
        let b = doc.create_element_typed("button", "");
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
        let a = doc.create_element_typed("div", "");
        let b = doc.create_element_typed("div", "");
        let child = doc.create_element_typed("button", "");
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
        let panel = doc.create_element_typed("div", "");
        let field = doc.create_element_typed("input", "text");
        assert!(doc.append_child(panel, field), "nesting off-document");
        // Reachable and configurable while the whole subtree is detached.
        doc.set_value(field, "typed");
        assert_eq!(doc.value(field), "typed");
        doc.set_style_property(field, "width", "70px");
        assert_eq!(doc.computed_style_property(field, "width"), "70px");

        assert!(doc.append_child(DOCUMENT, panel));
        assert!(
            doc.is_connected(field),
            "a child of an inserted node is connected"
        );
        assert_eq!(doc.value(field), "typed", "state survives insertion");
    }

    #[test]
    fn a_child_moves_out_of_a_nested_container() {
        let mut doc = Document::new("t");
        let outer = doc.create_element_typed("div", "");
        let inner = doc.create_element_typed("div", "");
        let b = doc.create_element_typed("button", "");
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
    fn a_leaf_takes_inline_content_and_refuses_a_box() {
        // Was `appending_to_a_leaf_reports_failure`, asserting a `<button>`
        // refuses a `<span>`. **That assertion was wrong about HTML.**
        // `<button><span>7</span></button>` is valid markup, every browser
        // accepts it, and it is how Flutter spells every label
        // (`ElevatedButton(child: Text('7'))`) — refusing it rendered a whole
        // keypad as empty rectangles.
        //
        // The rule the original test was reaching for survives, and is now
        // stated precisely: INLINE content is a leaf's caption; a BOX is a
        // genuine structural error and is still refused rather than dropped.
        let mut doc = Document::new("t");
        let b = doc.create_element_typed("button", "");
        doc.append_child(DOCUMENT, b);

        let label = doc.create_element_typed("span", "");
        doc.set_text_content(label, "7");
        assert!(doc.append_child(b, label), "a leaf takes inline content");
        assert_eq!(doc.text_content(b), "7", "the span is the button's caption");

        // A block-level child is still not a caption, and still reports.
        let box_ = doc.create_element_typed("div", "");
        assert!(!doc.append_child(b, box_), "a leaf cannot take a box");
        assert!(!doc.is_connected(box_));
        // Still usable — not lost.
        doc.set_text_content(box_, "still here");
        assert_eq!(doc.text_content(box_), "still here");
    }

    #[test]
    fn a_node_cannot_contain_itself() {
        let mut doc = Document::new("t");
        let outer = doc.create_element_typed("div", "");
        let inner = doc.create_element_typed("div", "");
        doc.append_child(DOCUMENT, outer);
        doc.append_child(outer, inner);
        assert!(!doc.append_child(inner, outer), "cycles must be refused");
        assert!(doc.is_connected(inner), "the refused move changes nothing");
    }

    #[test]
    fn removing_a_nested_child_takes_it_out_of_the_document() {
        let mut doc = Document::new("t");
        let panel = doc.create_element_typed("div", "");
        let b = doc.create_element_typed("button", "");
        doc.append_child(DOCUMENT, panel);
        doc.append_child(panel, b);
        assert!(doc.remove_child(panel, b));
        assert!(!doc.is_connected(b));
        // Re-insertable, because removeChild detaches rather than destroys.
        assert!(doc.append_child(DOCUMENT, b));
        assert!(doc.is_connected(b));
    }

    #[test]
    fn style_geometry_lands_on_the_control() {
        let mut doc = Document::new("t");
        let b = doc.create_element_typed("button", "");
        doc.append_child(DOCUMENT, b);
        doc.set_style_property(b, "left", "10px");
        doc.set_style_property(b, "top", "20px");
        doc.set_style_property(b, "width", "80px");
        assert_eq!(doc.computed_style_property(b, "left"), "10px");
        assert_eq!(doc.computed_style_property(b, "top"), "20px");
        assert_eq!(doc.computed_style_property(b, "width"), "80px");
    }

    #[test]
    fn a_row_of_inline_blocks_is_a_row() {
        // **The calculator's own tree.** `examples/js/calculator.js` appends a
        // `<div>` and four `<button>`s to it, sets no styles at all, and gets a
        // keypad — in a browser, because the UA stylesheet makes a button
        // `inline-block` and a div a block container running normal flow.
        //
        // Here it came out as four full-width bars stacked vertically: the div
        // was a `FlowLayoutPanel`, every container ran the FLEX algorithm, and
        // `align-items: stretch` sets each child's cross size to the
        // container's. The buttons' `inline-block` was computed correctly the
        // whole time and no layout ever read it.
        let mut doc = Document::new("t");
        let row = doc.create_element_typed("div", "");
        doc.append_child(DOCUMENT, row);
        let keys: Vec<NodeId> = ["7", "8", "9", "/"]
            .iter()
            .map(|label| {
                let key = doc.create_element_typed("button", "");
                let text = doc.create_text_node(label);
                doc.append_child(key, text);
                doc.append_child(row, key);
                key
            })
            .collect();

        let rects: Vec<LayoutRect> = keys.iter().map(|k| doc.border_rect(*k)).collect();
        // One line: every button shares a top edge.
        let top = rects[0].y;
        for (i, r) in rects.iter().enumerate() {
            assert!(
                (r.y - top).abs() < 0.5,
                "button {i} left the line: y={} vs {top}",
                r.y
            );
        }
        // Left to right, each starting where the last ended — no `spacing`,
        // because CSS puts nothing between adjacent boxes but their margins.
        for i in 1..rects.len() {
            let expected = rects[i - 1].x + rects[i - 1].w;
            assert!(
                (rects[i].x - expected).abs() < 0.5,
                "button {i} at x={}, expected {expected}",
                rects[i].x
            );
        }
        // And each keeps its own width rather than being stretched to the
        // container's — the half `align-items: stretch` was destroying.
        let container = doc.border_rect(row);
        for (i, r) in rects.iter().enumerate() {
            assert!(
                r.w < container.w / 2.0,
                "button {i} was stretched to {} of {}",
                r.w,
                container.w
            );
        }
    }

    #[test]
    fn an_author_rule_that_says_nothing_about_display_leaves_it_alone() {
        // A page styling its buttons must not stop them being inline-level.
        // The UA sheet says `button { display: inline-block }` and a `.key`
        // rule setting only a width has no opinion about display, so the
        // cascade has to keep the UA layer underneath it.
        let mut doc = Document::new("t");
        let sheet = doc.create_element_typed("style", "");
        let css = doc.create_text_node(".key { width: 60px; height: 40px; }");
        doc.append_child(sheet, css);
        doc.append_child(DOCUMENT, sheet);

        let row = doc.create_element_typed("div", "");
        doc.append_child(DOCUMENT, row);
        let keys: Vec<NodeId> = ["7", "8"]
            .iter()
            .map(|_| {
                let key = doc.create_element_typed("button", "");
                doc.set_attribute(key, "class", "key");
                doc.append_child(row, key);
                key
            })
            .collect();

        assert_eq!(
            doc.get_computed_style(keys[0]).display,
            Some(crate::css::Display::InlineBlock),
            "an author rule with no `display` dropped the UA layer's"
        );
        let (a, b) = (doc.border_rect(keys[0]), doc.border_rect(keys[1]));
        assert_eq!(a.y, b.y, "styled keys left the line box");
        assert!(a.w < 100.0, "a styled key was stretched to {}", a.w);
    }

    #[test]
    fn the_calculator_keypad_end_to_end() {
        // `examples/js/calculator.js` step for step, including the parts the
        // narrower tests leave out: a text-node child on each key, a class on
        // the row, and a stylesheet appended BEFORE anything it styles.
        let mut doc = Document::new("t");
        let sheet = doc.create_element_typed("style", "");
        let css = doc.create_text_node(
            ".row { margin-left: 8px; } .key { width: 60px; height: 56px; margin: 2px; }",
        );
        doc.append_child(sheet, css);
        doc.append_child(DOCUMENT, sheet);

        let row = doc.create_element_typed("div", "");
        doc.set_attribute(row, "class", "row");
        doc.append_child(DOCUMENT, row);

        let keys: Vec<NodeId> = ["7", "8", "9", "/"]
            .iter()
            .map(|label| {
                let key = doc.create_element_typed("button", "");
                doc.set_attribute(key, "class", "key");
                let text = doc.create_text_node(label);
                doc.append_child(key, text);
                doc.append_child(row, key);
                key
            })
            .collect();

        assert_eq!(
            doc.get_computed_style(keys[0]).display,
            Some(crate::css::Display::InlineBlock),
            "the cascade lost the UA layer's inline-block"
        );
        let rects: Vec<LayoutRect> = keys.iter().map(|k| doc.border_rect(*k)).collect();
        for (i, r) in rects.iter().enumerate() {
            assert_eq!(r.y, rects[0].y, "key {i} left the line box");
            assert!(
                (r.w - 60.0).abs() < 0.5,
                "key {i} is {} wide, not the 60px the sheet asked for",
                r.w
            );
        }
    }

    #[test]
    fn a_stylesheet_rule_can_position_and_paint() {
        // Two reports, one suspicion: a rule in a `<style>` reaches the CASCADE
        // but not the WIDGET. `position` and `background` are the two that were
        // called in, so both are asked here in one document.
        let mut doc = Document::new("t");
        let sheet = doc.create_element_typed("style", "");
        let css = doc.create_text_node(
            "#box { position: absolute; left: 40px; top: 25px; background: #ff0000; }",
        );
        doc.append_child(sheet, css);
        doc.append_child(DOCUMENT, sheet);

        let box_ = doc.create_element_typed("div", "");
        doc.set_attribute(box_, "id", "box");
        doc.append_child(DOCUMENT, box_);

        // The cascade's side — if this fails the selector never matched.
        assert_eq!(
            doc.get_computed_style(box_).position,
            Some(crate::css::Position::Absolute),
            "the rule never reached the cascade"
        );
        // …and the layout's side, which is the one being questioned.
        let rect = doc.border_rect(box_);
        assert_eq!(
            (rect.x, rect.y),
            (40.0, 25.0),
            "a stylesheet `position: absolute` did not place the box"
        );
        // `background` is the SHORTHAND, and it has to expand. Read from the
        // cascade, not from `style_property` — that reflects the element's own
        // declarations, and a sheet's value correctly is not among them.
        assert!(
            doc.get_computed_style(box_).background_color.is_some(),
            "the `background` shorthand did not expand to a colour"
        );
    }

    #[test]
    fn a_body_rule_paints_the_page() {
        // Reported from two sides: every other declaration in a sheet applies
        // and `body { background: … }` leaves the page its default grey.
        // `apply_style_property` HAS a `node == DOCUMENT` arm that paints the
        // form, so the question is whether a `body` selector ever reaches it.
        let mut doc = Document::new("t");
        let sheet = doc.create_element_typed("style", "");
        let css = doc.create_text_node("body { background: #202020; }");
        doc.append_child(sheet, css);
        doc.append_child(DOCUMENT, sheet);

        assert!(
            doc.get_computed_style(DOCUMENT).background_color.is_some(),
            "a `body` rule never reached the document's computed style"
        );
        assert_eq!(
            doc.page_background(),
            (32, 32, 32, 255),
            "the page kept its default ground"
        );
    }

    #[test]
    fn display_grid_places_items_in_its_track_template() {
        // `display: grid` parsed, cascaded, and ran NORMAL FLOW — only `flex`
        // selected a formatting context and everything else fell through. This
        // is the layout that was missing.
        let mut doc = Document::new("t");
        let sheet = doc.create_element_typed("style", "");
        let css = doc.create_text_node(
            "#g { display: grid; grid-template-columns: 100px 1fr; \
             width: 400px; column-gap: 20px; row-gap: 10px; }",
        );
        doc.append_child(sheet, css);
        doc.append_child(DOCUMENT, sheet);

        let grid = doc.create_element_typed("div", "");
        doc.set_attribute(grid, "id", "g");
        doc.append_child(DOCUMENT, grid);
        let cells: Vec<NodeId> = (0..4)
            .map(|_| {
                let cell = doc.create_element_typed("div", "");
                doc.append_child(grid, cell);
                cell
            })
            .collect();

        let r: Vec<LayoutRect> = cells.iter().map(|c| doc.border_rect(*c)).collect();
        // Two columns: the first fixed, the second taking what is left after
        // the fixed track and the gap.
        assert!(
            (r[0].w - 100.0).abs() < 0.5,
            "fixed track is {} wide, not 100",
            r[0].w
        );
        assert!(
            (r[1].x - (r[0].x + 100.0 + 20.0)).abs() < 0.5,
            "column-gap did not separate the tracks: {} vs {}",
            r[1].x,
            r[0].x
        );
        assert!(
            r[1].w > 200.0,
            "the `1fr` track did not take the leftover: {}",
            r[1].w
        );
        // Row-major auto-placement: the third item starts a new row.
        assert_eq!(r[0].y, r[1].y, "the first row split across two rows");
        assert!(r[2].y > r[0].y, "the third item did not start a new row");
        assert_eq!(r[2].x, r[0].x, "a new row did not return to the first track");
    }

    #[test]
    fn grid_placement_spans_tracks_and_flows_around_a_pinned_item() {
        // Auto-placement alone is a table. `grid-column: 1 / 3` is what makes a
        // grid a layout, and it had nowhere to land — the placement was
        // `slot % col_count` and nothing else.
        let mut doc = Document::new("t");
        let sheet = doc.create_element_typed("style", "");
        let css = doc.create_text_node(
            "#g { display: grid; grid-template-columns: 100px 100px 100px; \
             width: 300px; column-gap: 0; row-gap: 0; } \
             .wide { grid-column: 1 / 3; }",
        );
        doc.append_child(sheet, css);
        doc.append_child(DOCUMENT, sheet);

        let grid = doc.create_element_typed("div", "");
        doc.set_attribute(grid, "id", "g");
        doc.append_child(DOCUMENT, grid);

        // A wide item spanning two tracks, then three ordinary ones.
        let wide = doc.create_element_typed("div", "");
        doc.set_attribute(wide, "class", "wide");
        doc.append_child(grid, wide);
        let rest: Vec<NodeId> = (0..3)
            .map(|_| {
                let c = doc.create_element_typed("div", "");
                doc.append_child(grid, c);
                c
            })
            .collect();

        let w = doc.border_rect(wide);
        assert!(
            (w.w - 200.0).abs() < 0.5,
            "the spanning item is {} wide, not two 100px tracks",
            w.w
        );
        // The next item goes in the third column of the same row — the cursor
        // flowed AROUND the span rather than through it.
        let a = doc.border_rect(rest[0]);
        assert_eq!(a.y, w.y, "the next item left the first row");
        assert!(
            (a.x - (w.x + 200.0)).abs() < 0.5,
            "the next item overlapped the span: x={} vs span end {}",
            a.x,
            w.x + 200.0
        );
        // Row one is full, so the following item starts row two at column one.
        let b = doc.border_rect(rest[1]);
        assert!(b.y > w.y, "the fourth item did not start a new row");
        assert_eq!(b.x, w.x, "a new row did not return to the first track");
    }

    #[test]
    fn a_track_list_mixes_every_unit() {
        // px, %, fr and auto in ONE template, which is how real grids are
        // written. 600px container, 8px gaps, 4 tracks → 24px of gap.
        //   100px  +  10% (60)  +  auto  +  1fr
        // leaves the `fr` everything the others did not take.
        let mut doc = Document::new("t");
        let sheet = doc.create_element_typed("style", "");
        let css = doc.create_text_node(
            "#g { display: grid; grid-template-columns: 100px 10% auto 1fr; \
             width: 600px; padding: 0; column-gap: 8px; row-gap: 0; } \
             .m { width: 50px; }",
        );
        doc.append_child(sheet, css);
        doc.append_child(DOCUMENT, sheet);

        let grid = doc.create_element_typed("div", "");
        doc.set_attribute(grid, "id", "g");
        doc.append_child(DOCUMENT, grid);
        // Only the `auto` track's item declares a width — that is what `auto`
        // measures. Declaring one on the others would win over stretch (a
        // declared size beats the track), which is correct and would make this
        // test measure the items instead of the tracks.
        let cells: Vec<NodeId> = (0..4)
            .map(|i| {
                let c = doc.create_element_typed("div", "");
                if i == 2 {
                    doc.set_attribute(c, "class", "m");
                }
                doc.append_child(grid, c);
                c
            })
            .collect();

        let r: Vec<LayoutRect> = cells.iter().map(|c| doc.border_rect(*c)).collect();
        assert!((r[0].w - 100.0).abs() < 0.5, "px track: {}", r[0].w);
        assert!((r[1].w - 60.0).abs() < 0.5, "percent track: {}", r[1].w);
        // The `fr` takes what is left: 600 - 24 gaps - 100 - 60 - auto(50).
        assert!(
            (r[3].w - (600.0 - 24.0 - 100.0 - 60.0 - 50.0)).abs() < 1.0,
            "fr track took {} of the leftover",
            r[3].w
        );
        // Every track sits where the previous one ended, plus the gap.
        for i in 1..4 {
            assert!(
                (r[i].x - (r[i - 1].x + r[i - 1].w + 8.0)).abs() < 0.5,
                "track {i} is misplaced at {}",
                r[i].x
            );
        }
    }

    #[test]
    fn minmax_floors_a_track_and_still_flexes() {
        // `minmax(200px, 1fr)` — the responsive idiom. The floor holds when
        // there is not enough room to share, and the track still grows when
        // there is.
        let mut doc = Document::new("t");
        let sheet = doc.create_element_typed("style", "");
        let css = doc.create_text_node(
            "#g { display: grid; grid-template-columns: minmax(200px, 1fr) 1fr; \
             width: 300px; column-gap: 0; row-gap: 0; }",
        );
        doc.append_child(sheet, css);
        doc.append_child(DOCUMENT, sheet);
        let grid = doc.create_element_typed("div", "");
        doc.set_attribute(grid, "id", "g");
        doc.append_child(DOCUMENT, grid);
        let a = doc.create_element_typed("div", "");
        let b = doc.create_element_typed("div", "");
        doc.append_child(grid, a);
        doc.append_child(grid, b);

        // 300px between a floored track and a plain `1fr`: the floor wins and
        // the sibling gets the remainder.
        let ra = doc.border_rect(a);
        assert!(
            ra.w >= 199.5,
            "minmax floor was not honoured: {} < 200",
            ra.w
        );
        // Widen it and the floored track flexes above its floor.
        doc.set_style_property(grid, "width", "1000px");
        let wide = doc.border_rect(a).w;
        assert!(
            wide > ra.w,
            "the minmax track did not flex above its floor: {wide} vs {}",
            ra.w
        );
    }

    #[test]
    fn named_areas_place_items_by_name() {
        // `grid-template-areas` draws the layout; `grid-area` claims a room.
        // The template alone also sets the COLUMN COUNT — two names per row is
        // a two-column grid with no `grid-template-columns` in sight.
        let mut doc = Document::new("t");
        let sheet = doc.create_element_typed("style", "");
        let css = doc.create_text_node(
            "#g { display: grid; width: 400px; padding: 0; column-gap: 0; row-gap: 0; \
             grid-template-areas: \"head head\" \"side main\"; } \
             .h { grid-area: head; } .s { grid-area: side; } .m { grid-area: main; }",
        );
        doc.append_child(sheet, css);
        doc.append_child(DOCUMENT, sheet);
        let grid = doc.create_element_typed("div", "");
        doc.set_attribute(grid, "id", "g");
        doc.append_child(DOCUMENT, grid);

        #[allow(clippy::redundant_closure_call)]
        let mut cell = |class: &str| {
            let c = doc.create_element_typed("div", "");
            doc.set_attribute(c, "class", class);
            doc.append_child(grid, c);
            c
        };
        let (head, side, main) = (cell("h"), cell("s"), cell("m"));

        let (h, s, m) = (
            doc.border_rect(head),
            doc.border_rect(side),
            doc.border_rect(main),
        );
        // `head head` spans both columns of row one.
        assert!(
            (h.w - 400.0).abs() < 1.0,
            "the header did not span both columns: {}",
            h.w
        );
        // `side` and `main` share row two.
        assert_eq!(s.y, m.y, "side and main are not on the same row");
        assert!(s.y > h.y, "row two did not start below the header");
        assert!(m.x > s.x, "main is not to the right of side");
        assert!((s.w - 200.0).abs() < 1.0, "side is {} of a 2-col grid", s.w);
    }

    #[test]
    fn a_declared_width_survives_its_container() {
        // A block-level child with `width: 200px` was stretched to the
        // container anyway: the container was told each child's `display` and
        // never whether its width was the author's or `auto`.
        let mut doc = Document::new("t");
        let host = doc.create_element_typed("div", "");
        doc.append_child(DOCUMENT, host);
        let fixed = doc.create_element_typed("div", "");
        doc.append_child(host, fixed);
        doc.set_style_property(fixed, "width", "200px");
        let auto = doc.create_element_typed("div", "");
        doc.append_child(host, auto);

        assert!(
            (doc.border_rect(fixed).w - 200.0).abs() < 0.5,
            "a declared width was overruled: {}",
            doc.border_rect(fixed).w
        );
        assert!(
            doc.border_rect(auto).w > 300.0,
            "an `auto` width stopped filling: {}",
            doc.border_rect(auto).w
        );
    }

    #[test]
    fn a_div_paints_a_declared_background_and_border_and_nothing_otherwise() {
        // A `<div>` draws nothing by default — its background is `transparent`
        // and its border-style is `none` — and BOTH paint once declared. The
        // panel used to return from `paint` unless it was a `<fieldset>`, so a
        // declared background and a declared border were stored, consumed for
        // layout, and silently never drawn.
        let mut doc = Document::new("t");
        let plain = doc.create_element_typed("div", "");
        doc.append_child(DOCUMENT, plain);
        assert!(
            !doc.paints_own_box(plain),
            "an unstyled div must paint nothing"
        );

        let styled = doc.create_element_typed("div", "");
        doc.append_child(DOCUMENT, styled);
        doc.set_style_property(styled, "background", "#2d7ff9");
        assert!(
            doc.paints_own_box(styled),
            "a declared background did not reach the paint path"
        );

        // A width with no style is still `border-style: none`, so it paints
        // nothing — CSS's rule, and the reason the style is what is asked.
        let edged = doc.create_element_typed("div", "");
        doc.append_child(DOCUMENT, edged);
        doc.set_style_property(edged, "border-width", "3px");
        assert!(
            !doc.paints_own_border(edged),
            "a border-width with no style must not paint"
        );
        doc.set_style_property(edged, "border-style", "solid");
        assert!(
            doc.paints_own_border(edged),
            "a declared solid border did not reach the paint path"
        );
    }

    #[test]
    fn a_block_child_takes_a_row_of_its_own() {
        // The other half of normal flow, and the reason this is one algorithm
        // rather than an inline special case: a block box closes whatever line
        // is open and nothing shares its row.
        let mut doc = Document::new("t");
        let host = doc.create_element_typed("div", "");
        doc.append_child(DOCUMENT, host);
        let first = doc.create_element_typed("button", "");
        doc.append_child(host, first);
        let block = doc.create_element_typed("div", "");
        doc.append_child(host, block);
        let after = doc.create_element_typed("button", "");
        doc.append_child(host, after);

        let (a, b, c) = (
            doc.border_rect(first),
            doc.border_rect(block),
            doc.border_rect(after),
        );
        assert!(b.y >= a.y + a.h - 0.5, "the block shared the button's line");
        assert!(c.y >= b.y + b.h - 0.5, "the button ran alongside the block");
        // A block fills the content width; an inline-block does not.
        assert!(b.w > a.w, "the block box did not fill its container");
    }

    #[test]
    fn a_flex_container_still_runs_flex() {
        // The guard for Flutter. `Row`/`Column` reach `FlowLayoutPanel` by KIND
        // and never become elements, so the flex algorithm
        // has to keep behaving exactly as it did — normal flow is what a box
        // gets when the cascade did NOT say `flex`.
        let mut doc = Document::new("t");
        let row = doc.create_element_typed("div", "");
        doc.append_child(DOCUMENT, row);
        doc.set_style_property(row, "display", "flex");
        doc.set_style_property(row, "flex-direction", "row");
        // **A declared height, because that is what stretching is about.** With
        // `height: auto` the line is as tall as its tallest item and the item
        // then fills the line — both end up the item's own height, and the
        // assertion below could not tell stretching from not stretching. A row
        // with a height of its own is the case where the two differ, and it is
        // also every panel that reaches this widget by KIND, which is what this
        // guard is for. The auto-height rule has its own test in
        // `a_column_of_flex_rows_grows_to_hold_them`.
        doc.set_style_property(row, "height", "150px");
        let a = doc.create_element_typed("button", "");
        let b = doc.create_element_typed("button", "");
        doc.append_child(row, a);
        doc.append_child(row, b);
        // `align-items` defaults to `stretch`, so a flex child fills the cross
        // axis — the behaviour normal flow must NOT have and this must keep.
        let (ra, container) = (doc.border_rect(a), doc.border_rect(row));
        assert!(
            ra.h >= container.h - 8.5,
            "a flex child stopped stretching: {} of {}",
            ra.h,
            container.h
        );
    }

    #[test]
    fn a_style_the_write_side_accepts_reads_back() {
        // These three were WRITTEN (they issue widget commands) and read back
        // `""`, because the declaration was consumed into the command and
        // forgotten. Nothing recorded what the author actually said.
        let mut doc = Document::new("t");
        let b = doc.create_element_typed("button", "");
        doc.append_child(DOCUMENT, b);
        doc.set_style_property(b, "color", "red");
        doc.set_style_property(b, "background-color", "#fff");
        doc.set_style_property(b, "font-size", "18px");
        assert_eq!(doc.get_style_property(b, "color"), "red");
        assert_eq!(doc.get_style_property(b, "background-color"), "#fff");
        assert_eq!(doc.get_style_property(b, "font-size"), "18px");
    }

    #[test]
    fn a_style_nothing_renders_still_round_trips() {
        // The store accepts anything; layout consumes what it understands. An
        // unimplemented property is a rendering gap, not data loss.
        let mut doc = Document::new("t");
        let b = doc.create_element_typed("div", "");
        doc.append_child(DOCUMENT, b);
        doc.set_style_property(b, "mix-blend-mode", "multiply");
        assert_eq!(doc.get_style_property(b, "mix-blend-mode"), "multiply");
    }

    #[test]
    fn geometry_answers_from_the_control_not_the_declaration() {
        // A rect is COMPUTED. A docked element is not where its own `left`
        // said, so geometry must keep coming off the widget even though the
        // declaration is now recorded beside it.
        let mut doc = Document::new("t");
        let bar = doc.create_element_typed("div", "");
        doc.append_child(DOCUMENT, bar);
        doc.set_style_property(bar, "left", "300px");
        doc.set_style_property(bar, "dock", "top");
        assert_eq!(doc.computed_style_property(bar, "left"), "0px");
        // …while the declaration itself is intact underneath.
        assert_eq!(doc.style(bar).map(|s| s.get("left")), Some("300px"));
    }

    #[test]
    fn layout_can_read_the_declarations_it_will_need() {
        use crate::css::{Display, FlexDirection, Position};
        let mut doc = Document::new("t");
        let row = doc.create_element_typed("div", "");
        doc.append_child(DOCUMENT, row);
        doc.set_style_property(row, "display", "flex");
        doc.set_style_property(row, "flex-direction", "row");
        let props = doc.style_properties(row);
        assert!(props.is_flex_container());
        assert_eq!(props.display, Some(Display::Flex));
        assert_eq!(props.flex_direction, Some(FlexDirection::Row));

        let child = doc.create_element_typed("button", "");
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
        let b = doc.create_element_typed("button", "");
        doc.append_child(DOCUMENT, b);
        doc.set_style_property(b, "display", "none");
        assert!(doc.style_properties(b).display_none);
        assert_eq!(doc.get_style_property(b, "display"), "none");
    }

    /// A container sized so it actually runs a layout pass.
    ///
    /// `position: relative` is declared, because that is what makes it a
    /// containing block — a static box is transparent to its children's
    /// coordinates, in CSS and here. In a real program `primitives/gui.rs`
    /// emits this at construction for every container element; the fixture
    /// says it out loud for the same reason the markup does.
    /// A flex container of a given size.
    ///
    /// **`display: flex` is declared, and that is not a detail.** Every test
    /// built on this helper asserts flexbox — `align-items`, `align-self`,
    /// `order`, `justify-content`, `flex: 0`, children stacking in a column.
    /// None of them said `display: flex`, because at the time every container
    /// ran the flex algorithm whatever the cascade said. A `<div>` is
    /// `display: block` and now runs normal flow, so the declaration these
    /// tests always relied on has to be written down.
    ///
    /// Tests of normal flow deliberately do NOT use this helper — they build a
    /// bare `<div>`, which is what a page actually has.
    /// A flex COLUMN of the given size — the shape the tests below are about.
    ///
    /// ⚠ The axis is declared. This fixture used to say `display: flex` alone
    /// and get a column anyway, because the widget's own default was `TopDown`
    /// and an undeclared `flex-direction` never reached it. CSS says the
    /// initial value is `row` (css-flexbox §5.1), so that was a real defect —
    /// a `FlowLayoutPanel`, which is `display: flex; flex-wrap: wrap` and
    /// nothing else, stacked its buttons vertically instead of flowing them.
    ///
    /// Every test using this fixture is about `align-items`, `justify-content`,
    /// `order` or margins DOWN a column, and none of them is about which axis
    /// an undeclared container picks. Saying it out loud keeps them meaning
    /// what they were written to mean; the default itself is pinned by
    /// `a_flex_container_that_declares_no_axis_is_a_row`.
    fn container_with(doc: &mut Document, w: f32, h: f32) -> NodeId {
        let panel = doc.create_element_typed("div", "");
        doc.append_child(DOCUMENT, panel);
        doc.set_style_property(panel, "display", "flex");
        doc.set_style_property(panel, "flex-direction", "column");
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
        let b = doc.create_element_typed("button", "");
        doc.append_child(panel, b);
        doc.set_style_property(b, "left", "20px");
        doc.set_style_property(b, "top", "90px");
        doc.set_style_property(b, "width", "120px");
        assert_eq!(doc.computed_style_property(b, "left"), "20px");
        assert_eq!(doc.computed_style_property(b, "top"), "90px");
        assert_eq!(doc.computed_style_property(b, "width"), "120px");
    }

    #[test]
    fn coordinates_set_before_insertion_survive_joining_the_container() {
        // `Left := 8` before `Parent := Panel` is the ordinary way to write
        // VCL, and appending re-runs the container's layout.
        let mut doc = Document::new("t");
        let panel = container_with(&mut doc, 200.0, 200.0);
        let b = doc.create_element_typed("button", "");
        doc.set_style_property(b, "left", "30px");
        doc.set_style_property(b, "top", "40px");
        doc.append_child(panel, b);
        assert_eq!(doc.computed_style_property(b, "left"), "30px");
        assert_eq!(doc.computed_style_property(b, "top"), "40px");
    }

    #[test]
    fn a_child_with_no_coordinates_still_flows() {
        // Flutter never sets left/top — its children must keep being arranged,
        // or fixing the frontends that position by pixel breaks the one that
        // does not.
        let mut doc = Document::new("t");
        let panel = container_with(&mut doc, 200.0, 200.0);
        let a = doc.create_element_typed("button", "");
        let b = doc.create_element_typed("button", "");
        doc.append_child(panel, a);
        doc.append_child(panel, b);
        // Arranged top-down, so the second child sits below the first.
        let a_top = doc.computed_style_property(a, "top");
        let b_top = doc.computed_style_property(b, "top");
        assert_ne!(a_top, b_top, "flowed children must not overlap");
    }

    #[test]
    fn an_explicit_static_position_hands_the_child_back_to_the_container() {
        // The author overruling the inference: `left` is inert on a static box
        // in CSS, so saying `static` out loud means "you arrange me".
        let mut doc = Document::new("t");
        let panel = container_with(&mut doc, 200.0, 200.0);
        let b = doc.create_element_typed("button", "");
        doc.append_child(panel, b);
        doc.set_style_property(b, "position", "static");
        doc.set_style_property(b, "left", "60px");
        assert_ne!(doc.computed_style_property(b, "left"), "60px");
    }

    #[test]
    fn an_out_of_flow_child_takes_no_space_from_its_siblings() {
        // The half of `position: absolute` that is easy to miss: it is removed
        // from the flow, so the flowed sibling gets the whole container.
        let mut doc = Document::new("t");
        let panel = container_with(&mut doc, 200.0, 200.0);
        let flowed = doc.create_element_typed("button", "");
        doc.append_child(panel, flowed);
        let alone = doc.computed_style_property(flowed, "height");

        let positioned = doc.create_element_typed("button", "");
        doc.append_child(panel, positioned);
        doc.set_style_property(positioned, "left", "10px");
        doc.set_style_property(positioned, "top", "10px");
        assert_eq!(
            doc.computed_style_property(flowed, "height"),
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
            let n = doc.create_element_typed(tag, "");
            assert!(doc.append_child(DOCUMENT, n), "{tag} must be appendable");
            assert_eq!(doc.computed_style_property(n, "width"), "0px", "{tag} must not draw");
            assert_eq!(doc.computed_style_property(n, "height"), "0px");
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
        let list = doc.create_element_typed("select", "6");
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
        let list = doc.create_element_typed("select", "6");
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
        let group = doc.create_element_typed("fieldset", "");
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
        let tabs = doc.create_element_typed("vybe-tabcontrol", "");
        doc.append_child(DOCUMENT, tabs);
        // A tab control is not label-sized.
        assert_ne!(doc.computed_style_property(tabs, "width"), "120px");
        // …and it serialises as the custom element it is, so a browser build
        // knows exactly which `customElements.define` it needs.
        assert!(doc.to_html().contains("<vybe-tabcontrol>"));
    }

    #[test]
    fn a_custom_element_naming_no_control_degrades_visibly() {
        // A tag naming a control that does not exist is a mistake, and it must
        // fail toward something you can SEE rather than toward nothing.
        let mut doc = Document::new("t");
        let bogus = doc.create_element_typed("vybe-nonesuch", "");
        doc.append_child(DOCUMENT, bogus);
        assert_eq!(doc.computed_style_property(bogus, "width"), "120px");
    }

    #[test]
    fn a_non_visual_component_is_not_a_box() {
        // A Timer or ToolTip is a member of the form, not a rectangle on it —
        // WinForms and VCL both keep them in `components`, not `Controls`.
        // They drew as grey labels before.
        let mut doc = Document::new("t");
        for tag in ["vybe-timer", "vybe-tooltip", "vybe-imagelist"] {
            let n = doc.create_element_typed(tag, "");
            doc.append_child(DOCUMENT, n);
            assert_eq!(doc.computed_style_property(n, "width"), "0px", "{tag} must not draw");
        }
    }

    #[test]
    fn a_hidden_input_does_not_render() {
        let mut doc = Document::new("t");
        let hidden = doc.create_element_typed("input", "hidden");
        doc.append_child(DOCUMENT, hidden);
        assert_eq!(doc.computed_style_property(hidden, "width"), "0px");
        // …while an ordinary one does.
        let text = doc.create_element_typed("input", "text");
        doc.append_child(DOCUMENT, text);
        assert_ne!(doc.computed_style_property(text, "width"), "0px");
    }

    #[test]
    fn metadata_takes_no_slot_in_a_flow_container() {
        // A `<style>` in a container must not push its siblings down by a row
        // of nothing.
        let mut doc = Document::new("t");
        let panel = container_with(&mut doc, 200.0, 400.0);
        let first = doc.create_element_typed("button", "");
        doc.append_child(panel, first);
        let alone = doc.computed_style_property(first, "height");

        let style = doc.create_element_typed("style", "");
        doc.append_child(panel, style);
        assert_eq!(doc.computed_style_property(first, "height"), alone);
    }

    #[test]
    fn a_percentage_resolves_against_the_containing_block() {
        // Parsing cannot do this — `width: 50%` is a fraction of a parent the
        // parser cannot see, which is why `Length::Percent` stays symbolic
        // until something knows the containing block.
        let mut doc = Document::new("t");
        let panel = container_with(&mut doc, 400.0, 300.0);
        let b = doc.create_element_typed("button", "");
        doc.append_child(panel, b);
        doc.set_style_property(b, "position", "absolute");
        doc.set_style_property(b, "width", "50%");
        assert_eq!(doc.computed_style_property(b, "width"), "200px");
        doc.set_style_property(b, "height", "10%");
        assert_eq!(doc.computed_style_property(b, "height"), "30px");
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
            let panel = doc.create_element_typed("div", "");
            doc.append_child(DOCUMENT, panel);
            doc.set_style_property(panel, "position", "absolute");
            doc.set_style_property(panel, "width", "400px");
            doc.set_style_property(panel, "height", "300px");
            // `flex-direction` without `display: flex` is an inert
            // declaration — this test always meant a flex column and could
            // rely on getting one, because every container was flex.
            doc.set_style_property(panel, "display", "flex");
            doc.set_style_property(panel, "flex-direction", "column");

            let first = doc.create_element_typed("button", "");
            doc.append_child(panel, first);
            let second = doc.create_element_typed("button", "");
            doc.append_child(panel, second);

            doc.set_style_property(first, "position", mode);
            doc.set_style_property(first, "left", "25px");
            doc.set_style_property(first, "top", "10px");

            (
                doc.get_bounding_client_rect(first).expect("first is in the document"),
                doc.get_bounding_client_rect(second).expect("second is in the document"),
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

    /// Two stacked children that do NOT grow, the first carrying `margin`.
    /// Returns their tops.
    ///
    /// `flex: 0` is the point, not boilerplate. This container's children grow
    /// by default, and a growing child ABSORBS its own margin: the margin comes
    /// off the free space before it is shared out, so the box shrinks by as
    /// much as the margin added and the flow barely moves. That is correct
    /// flexbox and it makes margins unobservable — so a test about what a
    /// margin does to the flow has to hold the sizes still.
    fn stacked_with_margin(margin: &str) -> (f32, f32) {
        let mut doc = Document::new("t");
        let panel = container_with(&mut doc, 200.0, 300.0);
        let first = doc.create_element_typed("button", "");
        doc.append_child(panel, first);
        let second = doc.create_element_typed("button", "");
        doc.append_child(panel, second);
        doc.set_style_property(first, "flex", "0");
        doc.set_style_property(second, "flex", "0");
        if !margin.is_empty() {
            doc.set_style_property(first, "margin", margin);
        }
        (
            doc.get_bounding_client_rect(first).expect("first is in the document").y,
            doc.get_bounding_client_rect(second).expect("second is in the document").y,
        )
    }

    #[test]
    fn a_margin_pushes_the_element_and_everything_after_it() {
        // Margin is space BETWEEN siblings, so unlike a relative offset it
        // moves the rest of the flow too. That is the whole distinction, and
        // it is why the container owns margins rather than the child.
        let (plain_first, plain_second) = stacked_with_margin("");
        let (margin_first, margin_second) = stacked_with_margin("12px");

        assert_eq!(
            margin_first - plain_first,
            12.0,
            "the top margin pushes the element itself down"
        );
        assert_eq!(
            margin_second - plain_second,
            24.0,
            "and its siblings by the top AND bottom margin, since nothing \
             collapses here"
        );
    }

    #[test]
    fn an_out_of_flow_child_has_no_margin_to_claim() {
        // A margin is a claim on space between siblings. An absolutely
        // positioned box has no siblings in that sense — it was removed from
        // the flow — so its margin must not push anything.
        let mut doc = Document::new("t");
        let panel = container_with(&mut doc, 200.0, 300.0);
        let first = doc.create_element_typed("button", "");
        doc.append_child(panel, first);
        let second = doc.create_element_typed("button", "");
        doc.append_child(panel, second);

        let before = doc.get_bounding_client_rect(second).expect("second is in the document").y;
        doc.set_style_property(first, "position", "absolute");
        doc.set_style_property(first, "margin", "50px");
        let after = doc.get_bounding_client_rect(second).expect("second is in the document").y;

        assert!(
            after <= before,
            "an out-of-flow child's margin must not push its former siblings; \
             second moved from {before} to {after}"
        );
    }

    #[test]
    fn a_positioned_boxs_offsets_place_its_margin_edge_in_either_order() {
        // The other half of what a margin does. It claims no space between
        // siblings here — there are none — but it still displaces the box from
        // its own `left`/`top`, because the offsets place the MARGIN edge
        // (CSS 2.1 §10.3.7). Declaring the margin second must give the same
        // answer as declaring it first, which is the ordering bug.
        let place = |margin_first: bool| {
            let mut doc = Document::new("t");
            let panel = container_with(&mut doc, 400.0, 300.0);
            let box_ = doc.create_element_typed("div", "");
            doc.append_child(panel, box_);
            doc.set_style_property(box_, "position", "absolute");
            if margin_first {
                doc.set_style_property(box_, "margin", "5px");
            }
            doc.set_style_property(box_, "left", "20px");
            doc.set_style_property(box_, "top", "30px");
            if !margin_first {
                doc.set_style_property(box_, "margin", "5px");
            }
            let r = doc.get_bounding_client_rect(box_).expect("in the document");
            let cb = doc.containing_block(box_);
            (r.x - cb.x, r.y - cb.y)
        };

        assert_eq!(
            place(true),
            (25.0, 35.0),
            "left:20 + margin-left:5 puts the BORDER box at 25 from the block"
        );
        assert_eq!(
            place(false),
            place(true),
            "a margin declared after the offsets must move the box too"
        );
    }

    #[test]
    fn a_custom_property_inherits_and_a_var_reads_it() {
        // The point of custom properties: declare a theme once, high up, and
        // every control under it reads the same answer. The value is declared
        // on the container and consumed two levels down.
        let mut doc = Document::new("t");
        let panel = container_with(&mut doc, 400.0, 300.0);
        doc.set_style_property(panel, "--pad", "12px");

        let inner = doc.create_element_typed("div", "");
        doc.append_child(panel, inner);
        let leaf = doc.create_element_typed("button", "");
        doc.append_child(inner, leaf);
        doc.set_style_property(leaf, "padding", "var(--pad)");

        assert_eq!(
            doc.box_edges(leaf).padding.left,
            12.0,
            "the value comes from an ancestor two levels up"
        );
        // The STORE keeps what the author wrote — the CSSOM serialises the
        // specified value, not the substituted one.
        assert_eq!(
            doc.get_style_property(leaf, "padding"),
            "var(--pad)",
            "element.style must read back the var(), not its result"
        );
    }

    #[test]
    fn changing_a_custom_property_moves_what_already_read_it() {
        // The cost of inheritance, and the half that is easy to miss: the
        // declaration reading `--pad` was applied before `--pad` changed, so
        // something has to go back and re-apply it.
        let mut doc = Document::new("t");
        let panel = container_with(&mut doc, 400.0, 300.0);
        doc.set_style_property(panel, "--pad", "4px");
        let leaf = doc.create_element_typed("button", "");
        doc.append_child(panel, leaf);
        doc.set_style_property(leaf, "padding", "var(--pad)");
        assert_eq!(doc.box_edges(leaf).padding.left, 4.0);

        doc.set_style_property(panel, "--pad", "20px");
        assert_eq!(
            doc.box_edges(leaf).padding.left,
            20.0,
            "re-declaring the variable must move the box that read it"
        );
    }

    #[test]
    fn a_var_falls_back_and_an_unresolvable_one_drops_the_declaration() {
        let mut doc = Document::new("t");
        let panel = container_with(&mut doc, 400.0, 300.0);
        let leaf = doc.create_element_typed("button", "");
        doc.append_child(panel, leaf);

        doc.set_style_property(leaf, "padding", "var(--absent, 7px)");
        assert_eq!(
            doc.box_edges(leaf).padding.left,
            7.0,
            "an undeclared name takes the fallback"
        );

        // No value and no fallback is invalid at computed-value time: the
        // declaration is DROPPED, which is not the same as resolving to empty.
        doc.set_style_property(leaf, "padding", "var(--also-absent)");
        assert_eq!(
            doc.box_edges(leaf).padding.left,
            0.0,
            "an unresolvable reference leaves nothing behind"
        );
    }

    #[test]
    fn an_invalid_var_takes_the_box_back_rather_than_freezing_the_widget() {
        // The half `box_edges` cannot see. "Invalid at computed-value time"
        // means the element takes the value it would have had WITHOUT the
        // declaration — so the WIDGET has to be told, not just the store.
        // Asserted through a child's arranged offset, because that comes from
        // what the container was last told rather than from the cache.
        let mut doc = Document::new("t");
        let panel = container_with(&mut doc, 400.0, 300.0);
        let child = doc.create_element_typed("button", "");
        doc.append_child(panel, child);

        // The panel starts on its own intrinsic padding, which is NOT zero.
        let intrinsic = doc.get_bounding_client_rect(child).expect("in the document").x;
        assert!(intrinsic > 0.0, "the panel has a padding of its own");

        doc.set_style_property(panel, "padding", "10px");
        let padded = doc.get_bounding_client_rect(child).expect("in the document").x;
        assert!(padded > intrinsic, "the padding reached the container");

        // Now make the declaration unresolvable. The widget must be told
        // something — the bug this test exists for is it being told nothing and
        // sitting on `padded` forever.
        //
        // What it is told is CSS's answer, `padding: 0`, and NOT the panel's
        // intrinsic default: once a declaration has existed, the store is the
        // authority and the store says an undeclared padding is zero. So a
        // declared-then-invalidated padding does not restore the toolkit
        // default. That is a real divergence — recorded here rather than
        // hidden, like the margin-collapsing one — and it is still strictly
        // better than freezing.
        doc.set_style_property(panel, "padding", "var(--nope)");
        assert_eq!(
            doc.get_bounding_client_rect(child).expect("in the document").x,
            0.0,
            "an invalid declaration takes the CSS initial value, not the last one"
        );

        // The same door, and the one a theme actually opens: REMOVING a
        // variable that a live declaration reads.
        doc.set_style_property(panel, "--pad", "10px");
        doc.set_style_property(panel, "padding", "var(--pad)");
        assert_eq!(
            doc.get_bounding_client_rect(child).expect("in the document").x,
            padded,
            "read through a variable, the padding lands in the same place"
        );
        doc.set_style_property(panel, "--pad", "");
        assert_eq!(
            doc.get_bounding_client_rect(child).expect("in the document").x,
            0.0,
            "removing the variable must take the box back too"
        );
    }

    #[test]
    fn custom_property_names_are_case_sensitive_unlike_every_other_property() {
        // `--Brand` and `--brand` are two properties; `COLOR` and `color` are
        // one. Folding the custom name made a theme silently read the wrong
        // variable.
        let mut doc = Document::new("t");
        let panel = container_with(&mut doc, 400.0, 300.0);
        doc.set_style_property(panel, "--Gap", "30px");
        doc.set_style_property(panel, "--gap", "5px");

        let leaf = doc.create_element_typed("button", "");
        doc.append_child(panel, leaf);
        doc.set_style_property(leaf, "padding", "var(--Gap)");

        assert_eq!(
            doc.box_edges(leaf).padding.left,
            30.0,
            "--Gap and --gap are different properties"
        );
    }

    /// `measureText` answers in the font in effect, and says so when it
    /// cannot answer at all.
    ///
    /// The one canvas operation that ASKS. Two things are asserted: that the
    /// measurement tracks `setFont` (a wider face is a wider string, so the
    /// answer is not a fixed advance per character), and that a node with no
    /// surface answers `None` rather than `0.0` — an adapter has to be able to
    /// tell "no canvas" from "empty string".
    #[test]
    fn measuring_text_uses_the_font_in_effect_and_admits_when_it_cannot() {
        use crate::canvas::{Canvas as _, Font, FontStyle, FontWeight};

        let mut doc = Document::new("t");
        let canvas = doc.create_element_typed("canvas", "");
        doc.append_child(DOCUMENT, canvas);

        let small = doc.measure_text(canvas, "Hello").expect("a canvas measures");
        assert!(small > 0.0, "a non-empty string has width, got {small}");
        assert_eq!(doc.measure_text(canvas, ""), Some(0.0));

        doc.canvas_mut(canvas)
            .expect("the element owns a surface")
            .canvas_mut()
            .set_font(&Font {
                family: "sans-serif".to_string(),
                size: 48.0,
                weight: FontWeight::Normal,
                style: FontStyle::Normal,
            });
        let large = doc.measure_text(canvas, "Hello").expect("still a canvas");
        assert!(
            large > small * 2.0,
            "a 48px face is far wider than the 12px default: {small} then {large}"
        );

        // `save`/`restore` are paint state, and so is the font — a backwards
        // scan for the last `setFont` would return one that has been popped.
        doc.canvas_mut(canvas).unwrap().canvas_mut().save();
        doc.canvas_mut(canvas)
            .expect("the element owns a surface")
            .canvas_mut()
            .set_font(&Font {
                family: "sans-serif".to_string(),
                size: 8.0,
                weight: FontWeight::Normal,
                style: FontStyle::Normal,
            });
        doc.canvas_mut(canvas).unwrap().canvas_mut().restore();
        let restored = doc.measure_text(canvas, "Hello").expect("still a canvas");
        assert_eq!(
            restored, large,
            "restore puts back the 48px face, not the 8px one it replaced"
        );

        // A paragraph is not a drawing surface, and the absence is in the type.
        let p = doc.create_element_typed("p", "");
        doc.append_child(DOCUMENT, p);
        assert_eq!(doc.measure_text(p, "Hello"), None);
    }

    /// Writing an attribute on the DOCUMENT must not empty the body.
    ///
    /// `child_nodes` prefers a node's own `children` vec and only derives the
    /// document's from `order` when no `DomNode` exists for it — but
    /// `set_attribute` creates one with `entry(DOCUMENT).or_default()`, whose
    /// `children` is EMPTY. So one attribute write on the document made every
    /// body child invisible to flow, to `textContent`, and to the serialiser.
    #[test]
    fn an_attribute_on_the_document_does_not_orphan_the_body() {
        let mut doc = Document::new("t");
        let body = doc.body().expect("an HTML document has a body");
        let p = doc.create_element_typed("p", "");
        doc.append_child(DOCUMENT, p);
        // Appending to the document puts content in the BODY, as tree
        // construction does — the document's own child is `<html>`.
        assert_eq!(doc.child_nodes(body), vec![p]);

        doc.set_attribute(DOCUMENT, "lang", "en");
        assert_eq!(
            doc.child_nodes(body),
            vec![p],
            "the body still has its children after an attribute write on the document"
        );
        assert_eq!(
            doc.child_nodes(DOCUMENT).len(),
            1,
            "and the document keeps its single element child, `<html>`"
        );
    }

    /// `insertBefore` puts the child WHERE it was asked, at the root too.
    ///
    /// The document keeps its own child list now. Before that its children
    /// were derived by filtering creation order, so document order WAS
    /// creation order and this could not be expressed at all — an insert at
    /// the root silently appended.
    #[test]
    fn insert_before_orders_the_document_and_a_container_alike() {
        let mut doc = Document::new("t");
        let first = doc.create_element_typed("p", "");
        let last = doc.create_element_typed("p", "");
        doc.append_child(DOCUMENT, first);
        doc.append_child(DOCUMENT, last);

        let middle = doc.create_element_typed("p", "");
        assert!(doc.insert_before(DOCUMENT, middle, Some(last)));
        // The root's content lives in the body — `insertBefore` has to honour
        // the reference there just as it does in any other container.
        let body = doc.body().expect("an HTML document has a body");
        assert_eq!(doc.child_nodes(body), vec![first, middle, last]);

        // And inside a container, whose children live in its own `DomNode`.
        let box_ = doc.create_element_typed("div", "");
        doc.append_child(DOCUMENT, box_);
        let a = doc.create_element_typed("div", "");
        let c = doc.create_element_typed("div", "");
        doc.append_child(box_, a);
        doc.append_child(box_, c);
        let b = doc.create_element_typed("div", "");
        assert!(doc.insert_before(box_, b, Some(c)));
        assert_eq!(doc.child_nodes(box_), vec![a, b, c]);
    }

    // ── Intrinsic sizing ────────────────────────────────────────────────

    /// A paragraph is as tall as its text, and a longer text is taller.
    ///
    /// The whole of intrinsic sizing in one assertion: before it, every `<p>`
    /// was 200x20 whatever it said, so two paragraphs of wildly different
    /// length were the same box and everything after them sat in the wrong
    /// place. `default_size` said so itself — "a starting geometry, not a
    /// measurement".
    #[test]
    fn a_paragraph_is_as_tall_as_its_text() {
        let mut doc = Document::new("t");
        let short = doc.create_element_typed("p", "");
        doc.append_child(DOCUMENT, short);
        let text = doc.create_text_node("one line");
        doc.append_child(short, text);

        let long = doc.create_element_typed("p", "");
        doc.append_child(DOCUMENT, long);
        let words = doc.create_text_node(&"wrap ".repeat(200));
        doc.append_child(long, words);

        let short_h = doc.border_rect(short).h;
        let long_h = doc.border_rect(long).h;
        assert!(short_h > 0.0, "a line of text has a height");
        assert!(
            long_h > short_h * 4.0,
            "200 words wrap onto many lines: short={short_h}, long={long_h}"
        );
    }

    /// The paragraph fills the page's width, and its text wraps at that width
    /// rather than at a default the container never agreed to.
    ///
    /// The two halves are one fact: a height measured at the wrong width is
    /// the wrong height, so `width: auto` filling the containing block is what
    /// makes the measurement mean anything.
    #[test]
    fn a_block_fills_its_containing_block_and_wraps_there() {
        let mut doc = Document::new("t");
        let para = doc.create_element_typed("p", "");
        doc.append_child(DOCUMENT, para);
        let words = doc.create_text_node(&"wrap ".repeat(200));
        doc.append_child(para, words);

        let viewport = doc.form().rect().w;
        let rect = doc.border_rect(para);
        assert!(
            (rect.w - viewport).abs() < 0.5,
            "the paragraph fills the viewport: {} vs {viewport}",
            rect.w
        );

        // Narrow it and the same text needs more lines. This is the check that
        // the height came from the WRAPPING and not from a count of characters.
        let wide = rect.h;
        doc.set_style_property(para, "width", "200px");
        let narrow = doc.border_rect(para).h;
        assert!(
            narrow > wide,
            "narrower box, more lines: wide={wide}, narrow={narrow}"
        );
    }

    /// A declared height is the author's answer and content does not overrule
    /// it — which is also what keeps every pixel-positioned frontend still.
    #[test]
    fn a_declared_height_survives_its_content() {
        let mut doc = Document::new("t");
        let para = doc.create_element_typed("p", "");
        doc.append_child(DOCUMENT, para);
        doc.set_style_property(para, "height", "40px");
        let words = doc.create_text_node(&"overflowing ".repeat(200));
        doc.append_child(para, words);
        assert_eq!(doc.border_rect(para).h, 40.0);
    }

    /// **A pixel-positioned control is untouched by intrinsic sizing.**
    ///
    /// The check that matters for every existing frontend: a VCL or WinForms
    /// control carries `left`/`top`/`width`/`height`, which is `position:
    /// absolute` in all but the spelling — out of flow, and neither half of
    /// this applies to it. If either did, every Delphi window would reflow.
    #[test]
    fn a_positioned_control_keeps_the_geometry_it_declared() {
        let mut doc = Document::new("t");
        let label = doc.create_element_typed("label", "");
        doc.append_child(DOCUMENT, label);
        for (property, value) in [
            ("left", "40px"),
            ("top", "60px"),
            ("width", "120px"),
            ("height", "24px"),
        ] {
            doc.set_style_property(label, property, value);
        }
        let text = doc.create_text_node(&"a long caption ".repeat(20));
        doc.append_child(label, text);

        let rect = doc.border_rect(label);
        assert_eq!(
            (rect.x, rect.y, rect.w, rect.h),
            (40.0, 60.0, 120.0, 24.0),
            "its own coordinates, not the flow's"
        );
    }

    /// `button.appendChild(createTextNode("7"))` and `button.textContent = "7"`
    /// are the same thing — the spec defines the second AS the first.
    ///
    /// A leaf control has no line to lay a text node out on, so its text
    /// children are its caption. Refusing them made the two spellings differ:
    /// one drew "7" and the other drew an empty button, with nothing reported.
    #[test]
    fn a_leaf_controls_text_children_are_its_caption() {
        let mut doc = Document::new("t");
        let key = doc.create_element_typed("button", "");
        doc.append_child(DOCUMENT, key);
        let seven = doc.create_text_node("7");
        assert!(doc.append_child(key, seven), "a button accepts text");

        assert_eq!(doc.text_content(key), "7", "and answers it as its content");
        assert_eq!(doc.child_nodes(key), vec![seven], "as a NODE, not a string");

        // The other spelling, on a fresh button, reaches the same place.
        let other = doc.create_element_typed("button", "");
        doc.append_child(DOCUMENT, other);
        doc.set_text_content(other, "7");
        assert_eq!(doc.text_content(other), "7");
    }

    /// A box holding element children is not sized by its text alone.
    ///
    /// A block container with both kinds of child needs anonymous block boxes,
    /// which do not exist here — so sizing it to its inline runs would collapse
    /// its children out of view rather than approximate them.
    #[test]
    fn a_container_with_element_children_sizes_to_its_boxes_not_its_text() {
        // Was `…keeps_its_height`, asserting the container held its default
        // when it had element children. That was the boundary of intrinsic
        // sizing at the time, not the rule: the honest height of a block
        // container is where its FLOW finished, and now that its children are
        // laid out in normal flow that number exists.
        //
        // Stated against the child box rather than as "not the old value", so
        // it still fails if the height ever comes from the text runs again —
        // which is the mistake the original test was guarding against.
        let mut doc = Document::new("t");
        let host = doc.create_element_typed("div", "");
        doc.append_child(DOCUMENT, host);

        let inner = doc.create_element_typed("div", "");
        doc.append_child(host, inner);
        let text = doc.create_text_node("stray text");
        doc.append_child(host, text);

        let inner_h = doc.border_rect(inner).h;
        let host_h = doc.border_rect(host).h;
        assert!(
            host_h >= inner_h,
            "the container ({host_h}) is shorter than the box it holds \
             ({inner_h}) — sized to its text, not its boxes"
        );
    }

    /// A reference node that is not a child is `NotFoundError`, not an append.
    #[test]
    fn insert_before_a_stranger_is_refused_rather_than_appended() {
        let mut doc = Document::new("t");
        let host = doc.create_element_typed("div", "");
        doc.append_child(DOCUMENT, host);
        let stranger = doc.create_element_typed("p", "");
        doc.append_child(DOCUMENT, stranger);

        let child = doc.create_element_typed("span", "");
        assert!(
            !doc.insert_before(host, child, Some(stranger)),
            "the reference is not a child of the parent"
        );
        assert!(
            !doc.child_nodes(host).contains(&child),
            "and nothing was appended as a consolation"
        );
    }

    /// `replaceChild` swaps in place — the new node takes the old one's
    /// POSITION, which is the whole difference from remove-then-append.
    #[test]
    fn replace_child_keeps_the_position() {
        let mut doc = Document::new("t");
        let a = doc.create_element_typed("div", "");
        let b = doc.create_element_typed("div", "");
        let c = doc.create_element_typed("div", "");
        doc.append_child(DOCUMENT, a);
        doc.append_child(DOCUMENT, b);
        doc.append_child(DOCUMENT, c);

        let fresh = doc.create_element_typed("p", "");
        assert!(doc.replace_child(DOCUMENT, fresh, b));
        let body = doc.body().expect("an HTML document has a body");
        assert_eq!(doc.child_nodes(body), vec![a, fresh, c]);
    }

    /// `cloneNode(false)` copies the node and NOT its children; `true` copies
    /// the subtree. Neither puts the copy in the document.
    #[test]
    fn clone_node_is_shallow_until_it_is_deep() {
        let mut doc = Document::new("t");
        let host = doc.create_element_typed("div", "");
        doc.append_child(DOCUMENT, host);
        doc.set_attribute(host, "id", "original");
        doc.set_style_property(host, "color", "#ff0000");
        let inner = doc.create_element_typed("p", "");
        doc.append_child(host, inner);

        let shallow = doc.clone_node(host, false).expect("a node clones");
        assert!(doc.child_nodes(shallow).is_empty(), "shallow takes no children");
        assert_eq!(doc.get_attribute(shallow, "id").as_deref(), Some("original"));
        // The declarations came with it — a clone that lost its style would
        // look like a different element the moment it was inserted.
        assert_eq!(doc.get_computed_style(shallow).color, Some(0xffff0000));
        assert_eq!(doc.node(shallow).and_then(|n| n.parent), None, "not in the document");

        let deep = doc.clone_node(host, true).expect("a node clones");
        assert_eq!(doc.child_nodes(deep).len(), 1, "deep takes the subtree");
    }

    /// CDATA is a `Text`; a processing instruction is not.
    ///
    /// DOM §4.10 makes `CDATASection` a subclass of `Text`, so its data counts
    /// towards an ancestor's `textContent` exactly as a run does. A PI's does
    /// not — and both still answer their OWN data when asked directly. The two
    /// rules disagree on purpose and this pins which is which.
    #[test]
    fn cdata_counts_as_text_and_a_processing_instruction_does_not() {
        let mut doc = Document::new("t");
        let host = doc.create_element_typed("p", "");
        doc.append_child(DOCUMENT, host);

        let text = doc.create_text_node("plain ");
        let cdata = doc.create_cdata_section("raw<&>");
        let pi = doc.create_processing_instruction("xml-stylesheet", "href=\"x\"");
        let comment = doc.create_comment("invisible");
        assert!(doc.append_child(host, text));
        assert!(doc.append_child(host, cdata));
        assert!(doc.append_child(host, pi));
        assert!(doc.append_child(host, comment));

        assert_eq!(doc.text_content(host), "plain raw<&>");
        // Directly, every one of them is its own data.
        assert_eq!(doc.text_content(pi), "href=\"x\"");
        assert_eq!(doc.text_content(comment), "invisible");
        assert_eq!(doc.text_content(cdata), "raw<&>");
    }

    /// The three data kinds attach ANYWHERE — they place no requirement on the
    /// parent, because they contribute no line.
    ///
    /// A text node does place one, and a leaf meets it differently: it has no
    /// line, so the text becomes its caption rather than a run. It used to be
    /// refused outright, which is what made `appendChild` and `textContent`
    /// disagree — see `a_leaf_controls_text_children_are_its_caption`.
    #[test]
    fn data_nodes_need_no_line_capable_parent() {
        let mut doc = Document::new("t");
        // A `<label>` is a leaf: its text is one string, not a line of runs.
        let leaf = doc.create_element_typed("label", "");
        doc.append_child(DOCUMENT, leaf);

        let text = doc.create_text_node("caption");
        assert!(doc.append_child(leaf, text), "a leaf takes it as its text");
        assert_eq!(doc.text_content(leaf), "caption");

        for node in [
            doc.create_comment("c"),
            doc.create_cdata_section("d"),
            doc.create_processing_instruction("t", "v"),
        ] {
            assert!(
                doc.append_child(leaf, node),
                "data nodes attach regardless of the parent"
            );
        }
    }

    /// And each writes back as itself, unescaped — they are raw runs.
    #[test]
    fn the_xml_node_kinds_serialise_as_themselves() {
        let mut doc = Document::new("t");
        let host = doc.create_element_typed("div", "");
        doc.append_child(DOCUMENT, host);
        let cdata = doc.create_cdata_section("a < b & c");
        let pi = doc.create_processing_instruction("php", "echo 1;");
        let bare = doc.create_processing_instruction("target", "");
        doc.append_child(host, cdata);
        doc.append_child(host, pi);
        doc.append_child(host, bare);

        let html = doc.to_html();
        assert!(html.contains("<![CDATA[a < b & c]]>"), "got {html}");
        assert!(html.contains("<?php echo 1;?>"), "got {html}");
        assert!(html.contains("<?target?>"), "got {html}");
    }

    /// XML keeps its case; HTML folds it. The document decides, not the call.
    ///
    /// `create_element` folded unconditionally, because HTML tag names are
    /// case-insensitive and `VOID_ELEMENTS`, `control_kind` and the UA sheet
    /// all compare against lowercase literals. XML is case-SENSITIVE, so
    /// folding an XML document's names would turn `<Title>` into `<title>` and
    /// break `getElementsByTagName("Title")` for every caller that reads one.
    #[test]
    fn an_xml_document_keeps_the_case_an_html_one_folds() {
        let mut html = Document::new("t");
        let folded = html.create_element_typed("DIV", "");
        assert_eq!(html.node_name(folded), "div");
        html.set_attribute(folded, "DataRole", "x");
        assert_eq!(html.get_attribute(folded, "datarole").as_deref(), Some("x"));
        assert_eq!(html.get_elements_by_tag_name("Div"), vec![folded]);

        let mut xml = Document::new_xml("t");
        let kept = xml.create_element_typed("Title", "");
        assert_eq!(xml.node_name(kept), "Title");
        xml.set_attribute(kept, "DataRole", "x");
        // Case-sensitive both ways: the exact name finds it, a folded one does not.
        assert_eq!(xml.get_attribute(kept, "DataRole").as_deref(), Some("x"));
        assert_eq!(xml.get_attribute(kept, "datarole"), None);
        assert_eq!(xml.get_elements_by_tag_name("Title"), vec![kept]);
        assert!(xml.get_elements_by_tag_name("title").is_empty());
    }

    /// A namespaced element knows its vocabulary, and `prefix`/`localName`
    /// are two views of ONE name rather than two more fields.
    #[test]
    fn namespaces_are_recorded_and_the_qualified_name_is_split_not_stored() {
        let mut doc = Document::new_xml("t");
        let node = doc.create_element_ns(
            "http://www.w3.org/1999/XSL/Transform",
            "xsl:template",
            "",
        );
        assert_eq!(
            doc.namespace_uri(node).as_deref(),
            Some("http://www.w3.org/1999/XSL/Transform")
        );
        assert_eq!(doc.node_name(node), "xsl:template", "nodeName is QUALIFIED");
        assert_eq!(doc.prefix(node).as_deref(), Some("xsl"));
        assert_eq!(doc.local_name(node), "template");

        // No prefix is `None`, not `""`, and the local name is the whole tag.
        let bare = doc.create_element_ns("urn:x", "item", "");
        assert_eq!(doc.prefix(bare), None);
        assert_eq!(doc.local_name(bare), "item");

        // An element created without one has no namespace — `null`, per spec.
        let plain = doc.create_element_typed("item", "");
        assert_eq!(doc.namespace_uri(plain), None);

        // And a clone stays in its vocabulary.
        let copy = doc.clone_node(node, false).expect("elements clone");
        assert_eq!(
            doc.namespace_uri(copy).as_deref(),
            Some("http://www.w3.org/1999/XSL/Transform")
        );
    }

    /// Two attributes with the same local name in different vocabularies are
    /// TWO attributes — the thing a `HashMap<String, String>` could not hold.
    #[test]
    fn namespaced_attributes_do_not_collapse_onto_their_local_name() {
        let mut doc = Document::new_xml("t");
        let node = doc.create_element_typed("use", "");
        doc.append_child(DOCUMENT, node);

        doc.set_attribute(node, "href", "#plain");
        doc.set_attribute_ns(node, "http://www.w3.org/1999/xlink", "xlink:href", "#linked");

        // `getAttribute` matches the QUALIFIED name.
        assert_eq!(doc.get_attribute(node, "href").as_deref(), Some("#plain"));
        assert_eq!(
            doc.get_attribute(node, "xlink:href").as_deref(),
            Some("#linked")
        );
        // `getAttributeNS` matches namespace + LOCAL name — a different
        // question, and the reason both survive.
        assert_eq!(
            doc.get_attribute_ns(node, "http://www.w3.org/1999/xlink", "href")
                .as_deref(),
            Some("#linked")
        );
        assert_eq!(
            doc.get_attribute_ns(node, "", "href").as_deref(),
            Some("#plain")
        );
        assert_eq!(doc.node(node).map(|n| n.attributes.len()), Some(2));
    }

    /// Re-setting an attribute keeps its position — an attribute list is
    /// ORDERED, and a write is not a remove-then-append.
    #[test]
    fn setting_an_attribute_twice_does_not_move_it() {
        let mut doc = Document::new("t");
        let node = doc.create_element_typed("div", "");
        doc.append_child(DOCUMENT, node);
        doc.set_attribute(node, "data-a", "1");
        doc.set_attribute(node, "data-b", "2");
        doc.set_attribute(node, "data-a", "3");

        let names: Vec<String> = doc
            .node(node)
            .map(|n| n.attributes.iter().map(|a| a.name.clone()).collect())
            .unwrap_or_default();
        assert_eq!(names, vec!["data-a".to_string(), "data-b".to_string()]);
        assert_eq!(doc.get_attribute(node, "data-a").as_deref(), Some("3"));
    }

    /// One store, asked three ways.
    ///
    /// `element.style`, `getAttribute("style")` and a `[style]` selector are
    /// the same declarations (CSSOM §6.1). Writing the attribute has to reach
    /// the store, and both readers have to find it there — the attribute map
    /// no longer holds a copy, which is what made this worth asserting.
    #[test]
    fn the_style_attribute_and_the_declaration_store_are_one_thing() {
        let mut doc = Document::new("t");
        let p = doc.create_element_typed("p", "");
        doc.append_child(DOCUMENT, p);
        doc.set_attribute(p, "style", "color: #ff0000; padding: 4px");

        assert_eq!(doc.get_computed_style(p).color, Some(0xffff0000));
        let read_back = doc.get_attribute(p, "style").unwrap_or_default();
        assert!(read_back.contains("color"), "got {read_back:?}");
        assert_eq!(doc.query_selector_all("[style]"), vec![p]);
        assert!(
            doc.query_selector("p[style*=\"color\"]").is_some(),
            "a substring match reaches the store too"
        );

        // And an element nobody styled has no `style` attribute rather than an
        // empty one — presence is the whole of what `[style]` tests.
        let bare = doc.create_element_typed("p", "");
        doc.append_child(DOCUMENT, bare);
        assert_eq!(doc.get_attribute(bare, "style"), None);
        assert_eq!(doc.query_selector_all("[style]"), vec![p]);
    }

    #[test]
    fn a_link_is_blue_and_underlined_from_the_stylesheet() {
        // The UA layer's whole point: `<a>` arrives with a tag and nothing
        // else, and a rule is what makes it look like a link. Asserted through
        // the DOM rather than against `declarations_for`, because a rule that
        // parses and reaches no widget is the failure mode this replaced.
        let mut doc = Document::new("t");
        let panel = container_with(&mut doc, 200.0, 100.0);
        let link = doc.create_element_typed("a", "");
        doc.append_child(panel, link);

        // Compared against an author declaration of the same colour rather than
        // against a raw channel order, which is `css.rs`'s business.
        let reference = doc.create_element_typed("span", "");
        doc.append_child(panel, reference);
        doc.set_style_property(reference, "color", "#0000ee");

        let link_props = doc.style_properties(link);
        assert!(link_props.color.is_some(), "a link takes the UA colour");
        assert_eq!(
            link_props.color,
            doc.style_properties(reference).color,
            "and it is the spec's #0000ee"
        );
        assert_eq!(
            link_props.underline,
            Some(true),
            "and it is underlined"
        );
        // And it is a UA declaration, so it must not read back as inline style.
        assert_eq!(
            doc.style(link).map(|s| s.css_text()).unwrap_or_default(),
            "",
            "an unstyled <a> has an EMPTY element.style"
        );
    }

    #[test]
    fn an_inferred_positioned_box_takes_the_margin_too() {
        // The corpus case, and the one the two tests above could not see. Only
        // CONTAINERS are emitted with `position: absolute`; an ordinary VCL
        // control arrives with `left`/`top` and nothing else, and is inferred
        // out of flow. Gating the displacement on the DECLARATION made the
        // branch dead for every control on screen while the explicit tests
        // stayed green.
        let place = |margin: &str| {
            let mut doc = Document::new("t");
            let panel = container_with(&mut doc, 400.0, 300.0);
            let control = doc.create_element_typed("button", "");
            doc.append_child(panel, control);
            if !margin.is_empty() {
                doc.set_style_property(control, "margin-left", margin);
            }
            // No `position` — exactly what `emit_control_element` writes.
            doc.set_style_property(control, "left", "20px");
            doc.get_bounding_client_rect(control).expect("in the document").x
        };

        assert_eq!(
            place("7px") - place(""),
            7.0,
            "an inferred-positioned control is displaced by its margin"
        );
    }

    #[test]
    fn a_margin_does_not_displace_a_box_that_is_still_in_flow() {
        // The guard on the rule above. For an in-flow box the container applies
        // the margin, and adding it at the write site as well would count it
        // twice — which is why the offset arm asks about `position` at all.
        // Measured as a DELTA: a flow container places its children at its own
        // origin, which is not the containing block's, so the absolute x says
        // nothing on its own. How far the margin moved the box is the claim.
        let place = |margin: &str| {
            let mut doc = Document::new("t");
            let panel = container_with(&mut doc, 400.0, 300.0);
            let box_ = doc.create_element_typed("div", "");
            doc.append_child(panel, box_);
            doc.set_style_property(box_, "position", "relative");
            if !margin.is_empty() {
                doc.set_style_property(box_, "margin-left", margin);
            }
            doc.set_style_property(box_, "left", "20px");
            doc.get_bounding_client_rect(box_).expect("in the document").x
        };

        assert_eq!(
            place("5px") - place(""),
            5.0,
            "a relative box is offset by its margin ONCE — the container's \
             doing — not again at the write site"
        );
    }

    #[test]
    fn adjacent_margins_do_not_collapse_and_that_is_a_known_divergence() {
        // CSS collapses adjacent vertical margins: 10px below one child and
        // 10px above the next is 10px in a browser, and 20px here. Collapsing
        // needs block formatting contexts, which this container is not.
        //
        // Recorded as a test rather than a comment so the day it IS implemented
        // this fails and says exactly what changed.
        let mut doc = Document::new("t");
        let panel = container_with(&mut doc, 200.0, 300.0);
        let first = doc.create_element_typed("button", "");
        doc.append_child(panel, first);
        let second = doc.create_element_typed("button", "");
        doc.append_child(panel, second);
        // Held still for the same reason as `stacked_with_margin` — a growing
        // child eats its own margin and there is nothing left to collapse.
        doc.set_style_property(first, "flex", "0");
        doc.set_style_property(second, "flex", "0");

        let baseline = doc.get_bounding_client_rect(second).expect("second is in the document").y;
        doc.set_style_property(first, "margin-bottom", "10px");
        doc.set_style_property(second, "margin-top", "10px");
        let gap = doc.get_bounding_client_rect(second).expect("second is in the document").y - baseline;

        assert_eq!(
            gap, 20.0,
            "the two margins ADD here; a browser would collapse them to 10"
        );
    }

    #[test]
    fn the_font_axis_reaches_every_text_bearing_control_not_just_labels() {
        // `font-size` used to be a `let font_size = 13.0` INSIDE `Button::render`
        // — not a field, so no declaration could reach a button at all, and
        // weight/family/slant had no channel anywhere but the label. One shared
        // `FontSpec::apply_command` now serves all of them, which is also what
        // stops five widgets each deciding for themselves whether `bold` is 700.
        for tag in ["span", "button", "input", "textarea"] {
            let mut doc = Document::new("t");
            let node = doc.create_element_typed(tag, "");
            doc.append_child(DOCUMENT, node);
            doc.set_style_property(node, "font-weight", "bold");
            doc.set_style_property(node, "font-size", "20px");

            assert_eq!(
                doc.style_properties(node).font_weight,
                Some(700),
                "<{tag}> must take a declared font-weight"
            );
            assert_eq!(doc.get_style_property(node, "font-size"), "20px");
        }
    }

    #[test]
    fn a_ua_rule_reaches_the_control_without_becoming_an_inline_style() {
        // Both halves matter. A `<strong>` must actually be bold — a UA rule
        // that changes nothing is a table of no-ops — and `element.style` must
        // still be EMPTY, because a UA rule is not an inline style. Recording
        // it in the author store would serialise `style="font-weight:bold"`
        // onto every `<strong>` in the output.
        let mut doc = Document::new("t");
        let strong = doc.create_element_typed("strong", "");
        doc.append_child(DOCUMENT, strong);

        assert_eq!(doc.get_style_property(strong, "font-weight"), "");
        assert_eq!(
            doc.style_properties(strong).font_weight,
            Some(700),
            "the cascade must answer bold even though the author declared nothing"
        );
    }

    #[test]
    fn an_author_declaration_beats_the_user_agent() {
        // The cascade, in the only order that matters here: UA underneath,
        // author on top. Without it a `<strong>` could never be un-bolded.
        let mut doc = Document::new("t");
        let strong = doc.create_element_typed("strong", "");
        doc.append_child(DOCUMENT, strong);
        doc.set_style_property(strong, "font-weight", "normal");

        assert_eq!(doc.style_properties(strong).font_weight, Some(400));
        assert_eq!(doc.get_style_property(strong, "font-weight"), "normal");
    }

    #[test]
    fn a_block_box_keeps_its_text_and_its_children() {
        // The `<strong>` used to vanish: `<p>` mapped to a leaf label, a leaf
        // refuses `add_child`, and `append_child` put the child back in
        // `detached` — reported as `false`, which nobody checks, and rendered
        // as nothing.
        let mut doc = Document::new("t");
        let p = doc.create_element_typed("p", "");
        doc.append_child(DOCUMENT, p);
        doc.set_text_content(p, "a ");

        let strong = doc.create_element_typed("strong", "");
        assert!(
            doc.append_child(p, strong),
            "a paragraph must accept phrasing content"
        );
        assert_eq!(
            doc.text_content(p),
            "a ",
            "and keep its own text — a box has text AND children"
        );
        assert_eq!(
            doc.style_properties(strong).font_weight,
            Some(700),
            "the child is in the tree, so the UA sheet still reaches it"
        );
    }

    /// A table cell is a BOX, and an `<option>` still is not.
    ///
    /// Both halves are the point. `<td>` mapped to a leaf label on the
    /// reasoning that a cell is "a container's CONTENT" — the WinForms model.
    /// HTML says a cell holds flow content, so a `<div>` inside one was refused
    /// by `add_child` and silently went to `detached`. An `<option>` genuinely
    /// is the control's content and gets no box in a browser either, so it
    /// stays a leaf: this asserts a DECISION, not just a change.
    #[test]
    fn a_table_cell_holds_flow_content_but_an_option_does_not() {
        let mut doc = Document::new("t");
        let td = doc.create_element_typed("td", "");
        doc.append_child(DOCUMENT, td);
        doc.set_text_content(td, "cell ");

        let div = doc.create_element_typed("div", "");
        assert!(
            doc.append_child(td, div),
            "a cell must accept flow content — it can hold a div, a form, or another table"
        );
        assert_eq!(
            doc.text_content(td),
            "cell ",
            "and keep its own text: a box has text AND children"
        );

        let select = doc.create_element_typed("select", "");
        doc.append_child(DOCUMENT, select);
        let option = doc.create_element_typed("option", "");
        doc.append_child(select, option);
        let stray = doc.create_element_typed("div", "");
        assert!(
            !doc.append_child(option, stray),
            "an option is the control's content, not a box — it stays a leaf"
        );
    }

    /// `<strong>bold <em>and italic</em></strong>` — an inline element inside
    /// an inline element, which ordinary prose is full of.
    ///
    /// Inline elements have no `control_kind` arm and fall to `_ => "label"`,
    /// so this asks whether the silent-label trap reaches them too, or whether
    /// inline content escapes it by being RUNS built from the DOM rather than
    /// widgets built from the tree.
    #[test]
    fn an_inline_element_nests_inside_another_inline_element() {
        let mut doc = Document::new("t");
        let p = doc.create_element_typed("p", "");
        doc.append_child(DOCUMENT, p);

        let strong = doc.create_element_typed("strong", "");
        assert!(doc.append_child(p, strong), "a paragraph takes phrasing content");
        doc.set_text_content(strong, "bold ");

        let em = doc.create_element_typed("em", "");
        assert!(
            doc.append_child(strong, em),
            "phrasing content nests — `<strong>` must accept an `<em>`"
        );
        doc.set_text_content(em, "and italic");

        assert_eq!(
            doc.style_properties(em).font_weight,
            Some(700),
            "the `<em>` is inside the `<strong>`, so bold inherits to it"
        );
    }

    /// A link is TEXT, and still clickable.
    ///
    /// `<a>` was a link-label widget: a box parked on the line, unable to wrap
    /// and unable to hold the `<div>` HTML5's transparent content model allows.
    /// Making it a run fixes the flow — and would have silently deleted the
    /// click, which is why the exclusion stood. Both halves are asserted here
    /// because either alone is a regression.
    #[test]
    fn a_link_is_a_run_of_its_sentence_and_keeps_its_identity() {
        let mut doc = Document::new("t");
        let p = doc.create_element_typed("p", "");
        doc.append_child(DOCUMENT, p);
        doc.set_text_content(p, "see ");

        let a = doc.create_element_typed("a", "");
        doc.set_attribute(a, "href", "#x");
        assert!(doc.append_child(p, a), "a paragraph takes a link");
        doc.set_text_content(a, "the docs");

        // The UA sheet already dressed links; the widget was never what made
        // them blue.
        let props = doc.style_properties(a);
        assert_eq!(props.display, Some(crate::css::Display::Inline));
        assert!(
            props.color.is_some(),
            "the UA sheet colours a link — `a {{ color: #0000ee }}`"
        );

        // The sentence is ONE line: the box's own text plus the link's run,
        // not a paragraph with a control sitting in it.
        let runs = doc.inline_runs(p);
        assert_eq!(runs.len(), 2, "`see ` and `the docs` share the line: {runs:?}");
        assert_eq!(
            runs[0].source, None,
            "the box's own text belongs to no element"
        );
        assert!(
            runs[1].source.is_some(),
            "the link's run carries its source, which is what a click lands on"
        );
    }

    /// A POSITIONED link is still a `LinkLabel`, not a `Label`.
    ///
    /// The trap in making `<a>` inline: blockification sends a designer-placed
    /// link back through `control_kind`, and answering `label` there would have
    /// given it a widget whose `handle_mouse` returns `false` and which emits
    /// nothing — silently deleting the click and the pointer cursor from every
    /// WinForms and VCL LinkLabel. In flow it is a run; out of flow it is the
    /// control. Both, or neither is right.
    #[test]
    fn a_positioned_link_is_still_a_control_but_an_inline_one_is_a_run() {
        assert_eq!(
            control_kind("a", ""),
            "linklabel",
            "a blockified link needs a widget that can actually be clicked"
        );

        let mut doc = Document::new("t");
        let p = doc.create_element_typed("p", "");
        doc.append_child(DOCUMENT, p);

        let flowing = doc.create_element_typed("a", "");
        doc.append_child(p, flowing);
        doc.set_text_content(flowing, "in a sentence");
        assert!(
            doc.inline_runs(p).iter().any(|r| r.source.is_some()),
            "a link in flow is a run of its parent's line"
        );

        let placed = doc.create_element_typed("a", "");
        doc.append_child(DOCUMENT, placed);
        doc.set_style_property(placed, "position", "absolute");
        doc.set_style_property(placed, "left", "10px");
        doc.set_style_property(placed, "top", "20px");
        assert!(
            !doc.inline_runs(DOCUMENT)
                .iter()
                .any(|r| r.source.is_some()),
            "a positioned link left the flow — it is a box, not a run"
        );
    }

    /// A link in flow LOOKS clickable — the third thing after blue and
    /// underline, and the one a run could not have before.
    ///
    /// It arrives as a cascaded declaration, not as widget behaviour, which is
    /// the whole point: the painter never learns what a link is, so a `<span>`
    /// beside it keeps the ordinary cursor. Inherited, so an `<em>` inside the
    /// link is a hand too without a second rule.
    #[test]
    fn a_link_run_carries_the_pointer_cursor_and_plain_text_does_not() {
        let mut doc = Document::new("t");
        let p = doc.create_element_typed("p", "");
        doc.append_child(DOCUMENT, p);
        doc.set_text_content(p, "see ");

        let a = doc.create_element_typed("a", "");
        doc.append_child(p, a);
        doc.set_text_content(a, "the docs");

        let span = doc.create_element_typed("span", "");
        doc.append_child(p, span);
        doc.set_text_content(span, " and more");

        assert_eq!(
            doc.style_properties(a).cursor,
            Some(crate::css::Cursor::Pointer),
            "the UA sheet declares `a {{ cursor: pointer }}`"
        );
        assert_eq!(
            doc.style_properties(span).cursor,
            None,
            "ordinary phrasing content is not a hand"
        );

        let runs = doc.inline_runs(p);
        let link = runs
            .iter()
            .find(|r| r.text.contains("docs"))
            .expect("the link is a run");
        assert_eq!(
            link.cursor,
            Some(crate::css::Cursor::Pointer),
            "the run carries the resolved cursor, which is all the painter gets"
        );
    }

    /// The ROOT is a container too, and it was deaf.
    ///
    /// `document.body` IS the root node, so appending to it — the way every
    /// page built on the web platform is assembled — went through the one path
    /// that bailed instead of routing. The root never learned a child was
    /// inline-level, so it could only stack, and a form came out as a column
    /// down the left edge no matter what the cascade said.
    ///
    /// Two `<button>`s are the smallest case: `inline-block` in the UA sheet,
    /// so a browser puts them side by side.
    #[test]
    fn the_document_root_is_told_its_children_are_inline_level() {
        let mut doc = Document::new("t");
        doc.set_viewport(800.0, 600.0);

        let first = doc.create_element_typed("button", "");
        doc.append_child(DOCUMENT, first);
        doc.set_text_content(first, "One");

        let second = doc.create_element_typed("button", "");
        doc.append_child(DOCUMENT, second);
        doc.set_text_content(second, "Two");

        let (a, b) = (
            doc.get_bounding_client_rect(first).expect("in the document"),
            doc.get_bounding_client_rect(second).expect("in the document"),
        );
        assert_eq!(
            a.y, b.y,
            "inline-level siblings share a line box — they stacked: {a:?} vs {b:?}"
        );
        assert!(
            b.x >= a.x + a.w - 0.5,
            "the second button must follow the first across, not sit on it: {a:?} vs {b:?}"
        );
    }

    /// A field sits BESIDE its label, not underneath it.
    ///
    /// `<div><label>Name:</label><input></div>` is most of every form ever
    /// written. The `<label>` is inline, so it dissolves into a run of the
    /// div's text; the `<input>` is `inline-block`, so it keeps its box. Those
    /// are two different mechanisms, and until the widget joined the run
    /// sequence as an atomic slot they both started at the container's content
    /// origin — the label painted straight over the field.
    #[test]
    fn a_field_is_laid_out_after_its_label_not_on_top_of_it() {
        let mut doc = Document::new("t");
        doc.set_viewport(800.0, 600.0);
        let row = doc.create_element_typed("div", "");
        doc.append_child(DOCUMENT, row);

        let label = doc.create_element_typed("label", "");
        doc.append_child(row, label);
        doc.set_text_content(label, "A fairly long caption:");

        let field = doc.create_element_typed("input", "text");
        doc.append_child(row, field);

        let row_rect = doc.get_bounding_client_rect(row).expect("in the document");
        let field_rect = doc.get_bounding_client_rect(field).expect("in the document");
        assert!(
            field_rect.x > row_rect.x + 40.0,
            "the field must start after the caption's text, not at the row's \
             origin — row {row_rect:?}, field {field_rect:?}"
        );
    }

    /// `<progress value="60" max="100">` is six tenths full, not full.
    ///
    /// The widget stored a fraction and clamped every incoming value to
    /// `0.0..=1.0`, so the ordinary HTML spelling — a value in the units `max`
    /// declares — saturated on arrival and drew a complete bar. It looked
    /// correct for `value="0.6"` with no `max`, which is the one case a widget
    /// test would have written.
    #[test]
    fn a_progress_bar_is_a_fraction_of_its_max_not_of_one() {
        let mut doc = Document::new("t");
        let bar = doc.create_element_typed("progress", "");
        doc.append_child(DOCUMENT, bar);
        doc.set_attribute(bar, "max", "100");
        doc.set_attribute(bar, "value", "60");

        assert_eq!(
            doc.value(bar),
            "60",
            "the value reads back in the author's units, not renormalised"
        );

        // And the order the attributes arrive in must not change the answer —
        // a `max` after a `value` has to re-clamp rather than strand it.
        let other = doc.create_element_typed("progress", "");
        doc.append_child(DOCUMENT, other);
        doc.set_attribute(other, "value", "60");
        doc.set_attribute(other, "max", "100");
        assert_eq!(doc.value(other), "60", "max after value must not clamp it away");
    }

    /// `<input type="color">` is a picker, not a text field.
    ///
    /// `ColorPicker` was written, exported, `PanelWidget`-implementing and
    /// emitting `ColorChanged` — which `drain_events` already turned into
    /// `input`/`change`. Nothing named it, so the input fell to the text arm
    /// and rendered a box you could type prose into. The whole gap was one
    /// missing arm in each of the two halves of the mapping.
    ///
    /// This is also VB's `ColorDialog`, which is why the round-trip matters:
    /// the value has to read back as the `#rrggbb` HTML defines, not as a
    /// widget-shaped colour.
    #[test]
    fn a_colour_input_is_a_picker_and_round_trips_its_hex() {
        assert_eq!(control_kind("input", "color"), "colorpicker");

        let mut doc = Document::new("t");
        let picker = doc.create_element_typed("input", "color");
        doc.append_child(DOCUMENT, picker);
        doc.set_attribute(picker, "value", "#3366cc");

        assert_eq!(
            doc.value(picker),
            "#3366cc",
            "the picker holds the colour and answers it in HTML's spelling"
        );
    }

    /// Build `<table>` with one `<tr>` per row and the given cell texts.
    fn table_of(doc: &mut Document, rows: &[&[&str]]) -> (NodeId, Vec<Vec<NodeId>>) {
        let table = doc.create_element_typed("table", "");
        doc.append_child(DOCUMENT, table);
        let mut cells = Vec::new();
        for texts in rows {
            let row = doc.create_element_typed("tr", "");
            doc.append_child(table, row);
            let mut in_row = Vec::new();
            for text in *texts {
                let cell = doc.create_element_typed("td", "");
                doc.append_child(row, cell);
                doc.set_text_content(cell, text);
                in_row.push(cell);
            }
            cells.push(in_row);
        }
        (table, cells)
    }

    /// **What shape does the DOM actually build for a table?**
    ///
    /// The panel-level test proves a table with two rows sees two rows, and the
    /// DOM-level tests still place row two's cell nowhere — so the two layers
    /// disagree about the tree, and this asks the tree directly instead of
    /// inferring it from coordinates.
    #[test]
    fn a_table_widget_holds_rows_which_hold_cells() {
        let mut doc = Document::new("t");
        let (table, cells) = table_of(&mut doc, &[&["one", "two"], &["three", "four"]]);

        let name = Document::widget_name(table);
        let form = doc.form_mut();
        let table_widget = find_widget_mut(form, &name).expect("the table has a widget");
        let table_panel = table_widget
            .as_any_mut()
            .and_then(|any| any.downcast_mut::<crate::flow_layout::FlowLayoutPanel>())
            .expect("a table is a panel");

        let row_names: Vec<String> = table_panel
            .children_mut()
            .iter()
            .map(|c| c.name().to_string())
            .collect();
        assert_eq!(
            row_names.len(),
            2,
            "the table holds its two `<tr>`s, not their cells: {row_names:?}"
        );

        for (index, row) in table_panel.children_mut().iter_mut().enumerate() {
            let row_panel = row
                .as_any_mut()
                .and_then(|any| any.downcast_mut::<crate::flow_layout::FlowLayoutPanel>())
                .expect("a row is a panel");
            let held: Vec<String> = row_panel
                .children_mut()
                .iter()
                .map(|c| c.name().to_string())
                .collect();
            assert_eq!(
                held.len(),
                2,
                "row {index} holds its two cells: {held:?}"
            );
            for column in 0..2 {
                assert!(
                    held.contains(&Document::widget_name(cells[index][column])),
                    "row {index} holds the cell the DOM says it does: {held:?}"
                );
            }
        }
    }

    /// **A table lays its cells out in COLUMNS.**
    ///
    /// `<table>` used to answer the `datagridview` WIDGET — an HTML *layout*
    /// impersonating a .NET *control*. The `<tr>`s and `<td>`s inside it became
    /// generic panels beside a DataGrid that had never heard of them, so a
    /// table rendered as a grid control with its actual content stacked
    /// somewhere else.
    #[test]
    fn a_table_puts_the_cells_of_a_row_side_by_side() {
        let mut doc = Document::new("t");
        let (_, cells) = table_of(&mut doc, &[&["one", "two"]]);

        let a = doc.get_bounding_client_rect(cells[0][0]).expect("the first cell is laid out");
        let b = doc.get_bounding_client_rect(cells[0][1]).expect("the second cell is laid out");
        assert_eq!(a.y, b.y, "cells of one row share its top edge");
        assert!(
            b.x >= a.x + a.w,
            "the second cell starts after the first ends: {a:?} then {b:?}"
        );
    }

    /// Rows stack, and a column is the same x in every one of them — which is
    /// what makes it a column rather than two rows that happen to line up.
    #[test]
    fn a_column_is_the_same_x_in_every_row() {
        let mut doc = Document::new("t");
        let (_, cells) = table_of(&mut doc, &[&["one", "two"], &["three", "four"]]);

        let top = doc.get_bounding_client_rect(cells[0][0]).unwrap();
        let bottom = doc.get_bounding_client_rect(cells[1][0]).unwrap();
        assert!(
            bottom.y >= top.y + top.h,
            "the second row is below the first: {top:?} then {bottom:?}"
        );
        for column in 0..2 {
            let first = doc.get_bounding_client_rect(cells[0][column]).unwrap();
            let second = doc.get_bounding_client_rect(cells[1][column]).unwrap();
            assert_eq!(
                first.x, second.x,
                "column {column} is one column, not two coincidences"
            );
            assert_eq!(first.w, second.w, "and one width");
        }
    }

    /// **`<tfoot>` renders last however the markup was written** — §17.2.1.
    ///
    /// The whole reason the groups exist. Every other grouped test here appends
    /// the footer after the body, which is the one order that would pass with
    /// no reordering at all — so the reordering was untested by construction.
    ///
    /// This failed too, and for the same reason as the column-width test: the
    /// body row's cell was texted after the table's last measurement, so the
    /// row measured ~2px and the two rows overlapped. The ORDERING was right
    /// throughout — the footer sat one slot down, the slot was just empty.
    #[test]
    fn a_footer_written_first_still_renders_last() {
        let mut doc = Document::new("t");
        let table = doc.create_element_typed("table", "");
        doc.append_child(DOCUMENT, table);

        // The footer FIRST, which is legal and what §17.2.1 is about.
        let foot = doc.create_element_typed("tfoot", "");
        doc.append_child(table, foot);
        let foot_row = doc.create_element_typed("tr", "");
        doc.append_child(foot, foot_row);
        let total = doc.create_element_typed("td", "");
        doc.append_child(foot_row, total);
        doc.set_text_content(total, "total");

        let body = doc.create_element_typed("tbody", "");
        doc.append_child(table, body);
        let body_row = doc.create_element_typed("tr", "");
        doc.append_child(body, body_row);
        let value = doc.create_element_typed("td", "");
        doc.append_child(body_row, value);
        doc.set_text_content(value, "value");

        let footer = doc.get_bounding_client_rect(total).unwrap();
        let content = doc.get_bounding_client_rect(value).unwrap();
        assert!(
            footer.y >= content.y + content.h,
            "the footer is BELOW the body it was written above: {footer:?} vs {content:?}"
        );
    }

    /// **`table-layout: fixed` ignores the content** — §17.5.2.1.
    ///
    /// The whole difference from `auto`: the columns are decided from the
    /// table's width alone, so a column holding a long run is no wider than
    /// one holding a short one. Asserted against the SAME markup laid out both
    /// ways, because "equal widths" alone would also be true of a table whose
    /// contents happened to match.
    ///
    /// This caught a real defect on the `auto` half before it ever reached
    /// `fixed`: a single-row table split its width EVENLY and gave the wider
    /// column to the SHORTER text. Text reaches a `<td>` as a text node, and
    /// `insert_before`'s text-node branch returned without telling the
    /// enclosing table — so a cell's content never reached `column_widths` on
    /// the pass that used it. Two rows hid it, the other row having been
    /// measured in time to size the column.
    #[test]
    fn a_fixed_table_sizes_its_columns_without_asking_the_content() {
        let short = "id";
        let long = "a description long enough to want a column of its own";

        let mut auto = Document::new("t");
        let (_, auto_cells) = table_of(&mut auto, &[&[short, long]]);
        let auto_first = auto.get_bounding_client_rect(auto_cells[0][0]).unwrap();
        let auto_second = auto.get_bounding_client_rect(auto_cells[0][1]).unwrap();
        assert!(
            auto_second.w > auto_first.w,
            "auto gives the longer run the wider column: {} vs {}",
            auto_first.w,
            auto_second.w
        );

        let mut fixed = Document::new("t");
        let (fixed_table, fixed_cells) = table_of(&mut fixed, &[&[short, long]]);
        fixed.set_style_property(fixed_table, "table-layout", "fixed");
        let first = fixed.get_bounding_client_rect(fixed_cells[0][0]).unwrap();
        let second = fixed.get_bounding_client_rect(fixed_cells[0][1]).unwrap();
        assert!(
            (first.w - second.w).abs() < 1.0,
            "fixed divides the table evenly and never asks: {} vs {}",
            first.w,
            second.w
        );
    }

    /// **A cell with no row still gets one** — CSS 2.1 §17.2.1.
    ///
    /// The table generates an ANONYMOUS row box around cells that sit directly
    /// in it. Without that, `<table><th>a</th><th>b</th></table>` renders
    /// nothing at all, and a toolkit that appends CELLS to a grid — which is
    /// what `DataGridView.Columns.Add` does — can never produce a header.
    #[test]
    fn cells_directly_in_a_table_share_an_anonymous_row() {
        let mut doc = Document::new("t");
        let table = doc.create_element_typed("table", "");
        doc.append_child(DOCUMENT, table);

        let mut heads = Vec::new();
        for text in ["name", "email"] {
            let cell = doc.create_element_typed("th", "");
            doc.append_child(table, cell);
            doc.set_text_content(cell, text);
            heads.push(cell);
        }
        // …and a real row after them, which must NOT be swallowed into the
        // anonymous one.
        let row = doc.create_element_typed("tr", "");
        doc.append_child(table, row);
        let body_cell = doc.create_element_typed("td", "");
        doc.append_child(row, body_cell);
        doc.set_text_content(body_cell, "ada");

        let first = doc.get_bounding_client_rect(heads[0]).unwrap();
        let second = doc.get_bounding_client_rect(heads[1]).unwrap();
        assert!(
            second.x > first.x,
            "the loose cells are SIDE BY SIDE, which is what one row means: {first:?} then {second:?}"
        );
        assert_eq!(first.y, second.y, "and they share a top edge");

        let below = doc.get_bounding_client_rect(body_cell).unwrap();
        assert!(
            below.y >= first.y + first.h,
            "the real row follows the anonymous one: {first:?} then {below:?}"
        );
        assert_eq!(
            first.x, below.x,
            "and joins the column the loose cells established"
        );
    }

    /// **The table's BOX covers the rows it drew.**
    ///
    /// Every other test here asks where a cell went, and they all passed while
    /// a table's own border ran straight through its last row: the cells were
    /// placed correctly inside a rectangle measured before they existed. A
    /// table's height comes from its content, so the box has to follow the
    /// rows — nothing about a cell's position says whether it does.
    ///
    /// Found in a capture. Both directions were wrong at once: a table that
    /// gained a row clipped it, and one measured while it had more content
    /// kept the taller box and drew a large empty area under its last row.
    #[test]
    fn a_table_is_as_tall_as_its_rows() {
        let mut doc = Document::new("t");
        let (table, cells) = table_of(&mut doc, &[&["one", "two"], &["three", "four"]]);

        let box_of_table = doc.get_bounding_client_rect(table).unwrap();
        let last = doc.get_bounding_client_rect(cells[1][0]).unwrap();
        assert!(
            last.y + last.h <= box_of_table.y + box_of_table.h,
            "the last row is INSIDE the table's box: row {last:?} in {box_of_table:?}"
        );
        // And no further than it needs to be: a box much taller than its rows
        // draws an empty band under the table, which is the same defect from
        // the other side. One row's slack is spacing and borders, not content.
        let slack = (box_of_table.y + box_of_table.h) - (last.y + last.h);
        assert!(
            slack < last.h,
            "and the box does not trail an empty row's worth below it: {slack} spare"
        );
    }

    /// **A row added AFTER the table was laid out is still laid out.**
    ///
    /// Every other test here builds the whole tree and then asserts, which is
    /// the one order the invalidation already handled — a table whose last
    /// insertion happens to re-trigger it. This asks the question the other
    /// way round: the table has already been measured and placed once, and
    /// then it grows. A browser reflows; so must this.
    #[test]
    fn a_row_added_after_the_table_was_laid_out_is_placed() {
        let mut doc = Document::new("t");
        let (table, cells) = table_of(&mut doc, &[&["one", "two"]]);
        let first = doc.get_bounding_client_rect(cells[0][0]).unwrap();

        let row = doc.create_element_typed("tr", "");
        doc.append_child(table, row);
        let mut late = Vec::new();
        for text in ["three", "four"] {
            let cell = doc.create_element_typed("td", "");
            doc.append_child(row, cell);
            doc.set_text_content(cell, text);
            late.push(cell);
        }

        let below = doc.get_bounding_client_rect(late[0]).unwrap();
        assert!(
            below.y >= first.y + first.h,
            "the late row sits below the first: {first:?} then {below:?}"
        );
        assert_eq!(
            first.x, below.x,
            "and its cells join the columns that already existed"
        );
        let second = doc.get_bounding_client_rect(late[1]).unwrap();
        assert!(
            second.x > below.x,
            "its own cells are side by side too: {below:?} then {second:?}"
        );
    }

    /// **Restyling one cell re-measures the COLUMN.**
    ///
    /// A cell's padding is part of its min-content width, and column widths
    /// are computed from every cell in the column — so a style change on one
    /// cell is a table-wide question, not a cell-local one. Insertion and text
    /// changes route to the table; a style change has to as well, or the
    /// column keeps a width that no longer fits its contents.
    #[test]
    fn restyling_a_cell_re_measures_its_column() {
        let mut doc = Document::new("t");
        let (_, cells) = table_of(&mut doc, &[&["a", "b"], &["c", "d"]]);
        let before = doc.get_bounding_client_rect(cells[0][0]).unwrap();

        doc.set_attribute(cells[0][0], "style", "padding-left: 40px");

        let after = doc.get_bounding_client_rect(cells[0][0]).unwrap();
        assert!(
            after.w > before.w,
            "the padded cell got wider: {} then {}",
            before.w,
            after.w
        );
        let neighbour = doc.get_bounding_client_rect(cells[1][0]).unwrap();
        assert_eq!(
            after.w, neighbour.w,
            "and the cell BELOW it took the same width — a column has one width"
        );
        let right = doc.get_bounding_client_rect(cells[0][1]).unwrap();
        assert!(
            right.x >= after.x + after.w - 1.0,
            "the next column moved out of the way: {after:?} then {right:?}"
        );
    }

    /// `colspan` — the cell covers its neighbour's column, and the cell BELOW
    /// it still starts at column 0.
    ///
    /// This is the case a naive table gets wrong: a cell's column is not its
    /// index among its siblings once anything spans.
    #[test]
    fn a_colspan_cell_covers_the_columns_it_claims() {
        let mut doc = Document::new("t");
        let (_, cells) = table_of(&mut doc, &[&["wide"], &["a", "b"]]);
        doc.set_attribute(cells[0][0], "colspan", "2");

        let wide = doc.get_bounding_client_rect(cells[0][0]).unwrap();
        let left = doc.get_bounding_client_rect(cells[1][0]).unwrap();
        let right = doc.get_bounding_client_rect(cells[1][1]).unwrap();
        assert_eq!(wide.x, left.x, "the spanning cell starts at the first column");
        assert!(
            wide.w >= (right.x + right.w) - left.x - 1.0,
            "and reaches the end of the second: spans {} vs {}",
            wide.w,
            (right.x + right.w) - left.x
        );
    }

    /// `rowspan` — the slot a spanning cell occupies in the row BELOW is not
    /// available to that row, so its first cell is pushed into column 1.
    #[test]
    fn a_rowspan_cell_pushes_the_next_rows_cells_along() {
        let mut doc = Document::new("t");
        let (_, cells) = table_of(&mut doc, &[&["tall", "top"], &["bottom"]]);
        doc.set_attribute(cells[0][0], "rowspan", "2");

        let tall = doc.get_bounding_client_rect(cells[0][0]).unwrap();
        let top = doc.get_bounding_client_rect(cells[0][1]).unwrap();
        let bottom = doc.get_bounding_client_rect(cells[1][0]).unwrap();
        assert_eq!(
            bottom.x, top.x,
            "the second row's only cell sits in column 1, because column 0 is still occupied"
        );
        assert!(
            tall.h > top.h,
            "the spanning cell covers both rows: {} vs {}",
            tall.h,
            top.h
        );
    }

    /// A `<thead>` is not a row. Its ROWS are the table's rows — a table that
    /// treated the group as one row would give the whole header one row's
    /// height and lay its cells out as if they were the header's.
    #[test]
    fn a_row_group_contributes_its_rows_not_itself() {
        let mut doc = Document::new("t");
        let table = doc.create_element_typed("table", "");
        doc.append_child(DOCUMENT, table);
        let head = doc.create_element_typed("thead", "");
        doc.append_child(table, head);
        let row = doc.create_element_typed("tr", "");
        doc.append_child(head, row);
        let cell = doc.create_element_typed("td", "");
        doc.append_child(row, cell);
        doc.set_text_content(cell, "header");

        // A body row too, so the header has something to be a column WITH.
        let body = doc.create_element_typed("tbody", "");
        doc.append_child(table, body);
        let body_row = doc.create_element_typed("tr", "");
        doc.append_child(body, body_row);
        let body_cell = doc.create_element_typed("td", "");
        doc.append_child(body_row, body_cell);
        doc.set_text_content(body_cell, "value");

        // ⚠ NOT `w > 0 && h > 0` — that is what this test asserted at first,
        // and it passed while every grouped table rendered COMPLETELY EMPTY.
        // An unplaced cell keeps its construction size, so a nonzero box is
        // no evidence of layout at all. Only a POSITION the table chose is.
        let head_rect = doc.get_bounding_client_rect(cell).expect("a cell inside a thead is laid out");
        let body_rect = doc.get_bounding_client_rect(body_cell).unwrap();
        let table_rect = doc.get_bounding_client_rect(table).unwrap();
        assert!(
            head_rect.x > table_rect.x && head_rect.y > table_rect.y,
            "the header cell was placed INSIDE its table: {head_rect:?} in {table_rect:?}"
        );
        assert_eq!(
            head_rect.x, body_rect.x,
            "and it shares a column with the body's cell, across two groups"
        );
        assert!(
            body_rect.y >= head_rect.y + head_rect.h,
            "the body's row follows the header's: {head_rect:?} then {body_rect:?}"
        );
    }

    /// **The type arrives AFTER the element.**
    ///
    /// The DOM's `createElement` takes one argument, so a script has nowhere
    /// to say what kind of input it is building: `setAttribute("type", …)`
    /// afterwards is the only spelling there is. Nothing listened, and the
    /// element kept the text box it was born as — every scripted `<input>` in
    /// the engine, checkbox and colour and file alike, rendered as a field.
    ///
    /// Asserted on the CONTROL, never on `control_kind`: that mapping was
    /// right the whole time, which is exactly what hid this. Tests that built
    /// the element WITH its type — the only ones there were — all passed.
    #[test]
    fn a_type_set_after_creation_rebuilds_the_control() {
        let mut doc = Document::new("t");
        let input = doc.create_element_typed("input", "");
        doc.append_child(DOCUMENT, input);
        assert_eq!(
            doc.get_bounding_client_rect(input).map(|r| r.w),
            Some(160.0),
            "an input with no type is a text field — HTML's missing value default"
        );

        doc.set_attribute(input, "type", "color");
        assert_eq!(
            doc.get_bounding_client_rect(input).map(|r| r.w),
            Some(64.0),
            "the swatch is a different CONTROL, not a text box wearing a label"
        );
        doc.set_attribute(input, "value", "#3366cc");
        assert_eq!(doc.value(input), "#3366cc");

        // Removing the attribute is the missing value default, not "no type".
        doc.remove_attribute(input, "type");
        assert_eq!(
            doc.get_bounding_client_rect(input).map(|r| r.w),
            Some(160.0),
            "back to a text field, because that is what an input with no type is"
        );
    }

    /// Order must not matter. `value` then `type` and `type` then `value` are
    /// the same document, and the control is rebuilt in between — so anything
    /// the node already carries has to be replayed to the control that
    /// replaced the old one. This is `<progress>`'s lesson in a second place:
    /// there, `max` arriving after `value` clamped a value that was correct.
    #[test]
    fn attributes_written_before_the_type_survive_the_rebuild() {
        let mut doc = Document::new("t");
        let input = doc.create_element_typed("input", "");
        doc.append_child(DOCUMENT, input);

        doc.set_attribute(input, "value", "#3366cc");
        doc.set_attribute(input, "id", "swatch");
        doc.set_attribute(input, "type", "color");

        assert_eq!(
            doc.value(input),
            "#3366cc",
            "the value was written to the control that got replaced; the node still knew it"
        );
        assert_eq!(doc.get_attribute(input, "id").as_deref(), Some("swatch"));
    }

    /// A control the cascade already styled keeps that style across the swap.
    ///
    /// `restyle_subtree` cannot do this on its own: it applies what CHANGED,
    /// and nothing about the cascade changes when a widget is replaced — the
    /// new one has simply never been told any of it.
    #[test]
    fn a_rebuilt_control_is_told_the_style_it_already_had() {
        let mut doc = Document::new("t");
        let input = doc.create_element_typed("input", "");
        doc.append_child(DOCUMENT, input);
        doc.set_attribute(input, "style", "width: 40px");
        assert_eq!(doc.get_bounding_client_rect(input).map(|r| r.w), Some(40.0));

        doc.set_attribute(input, "type", "color");
        assert_eq!(
            doc.get_bounding_client_rect(input).map(|r| r.w),
            Some(40.0),
            "the declaration outranks the new kind's starting size, as it did the old one's"
        );
    }

    #[test]
    fn text_around_an_inline_element_survives_because_text_is_a_node() {
        // `a <b>B</b> c` — the case a single string could not hold. The
        // trailing ` c` had nowhere to live, so the markup was expressible and
        // the tree was not.
        let mut doc = Document::new("t");
        let p = doc.create_element_typed("p", "");
        doc.append_child(DOCUMENT, p);

        let before = doc.create_text_node("a ");
        doc.append_child(p, before);
        let strong = doc.create_element_typed("strong", "");
        doc.append_child(p, strong);
        doc.set_text_content(strong, "B");
        let after = doc.create_text_node(" c");
        doc.append_child(p, after);

        let runs = doc.inline_runs(p);
        assert_eq!(runs.len(), 3, "three runs, in document order");
        assert_eq!(runs[0].text, "a ");
        assert_eq!(runs[1].text, "B");
        assert_eq!(runs[2].text, " c");
        assert_eq!(
            runs[1].font.weight, 700,
            "only the middle run is bold — that IS the inline formatting context"
        );
        assert_eq!(runs[0].font.weight, 400);
        assert_eq!(
            doc.text_content(p),
            "a B c",
            "textContent is the concatenation of descendant text, per spec"
        );
    }

    #[test]
    fn source_formatting_collapses_the_way_html_says() {
        // Markup is written for humans: the newline and indent between tags
        // are formatting, not content. Rendering them verbatim is the most
        // visible way a renderer stops looking like HTML.
        let mut doc = Document::new("t");
        let p = doc.create_element_typed("p", "");
        doc.append_child(DOCUMENT, p);
        let lead = doc.create_text_node("\n    Hello   there\n    ");
        doc.append_child(p, lead);
        let strong = doc.create_element_typed("strong", "");
        doc.append_child(p, strong);
        doc.set_text_content(strong, "world");
        let tail = doc.create_text_node("\n    again\n");
        doc.append_child(p, tail);

        let runs = doc.inline_runs(p);
        assert_eq!(
            runs[0].text, "Hello there ",
            "indent dropped at the line start, the inner run of spaces becomes one, \
             and the space before <strong> survives"
        );
        assert_eq!(runs[1].text, "world");
        assert_eq!(
            runs[2].text, " again ",
            "the newline after </strong> is the space between the words"
        );
    }

    #[test]
    fn a_whitespace_only_run_still_separates_its_neighbours() {
        // `</b>\n  <b>` is one space, not nothing. Dropping it is how two
        // words the author wrote on separate lines end up jammed together.
        let mut doc = Document::new("t");
        let p = doc.create_element_typed("p", "");
        doc.append_child(DOCUMENT, p);
        let a = doc.create_element_typed("strong", "");
        doc.append_child(p, a);
        doc.set_text_content(a, "one");
        let gap = doc.create_text_node("\n   ");
        doc.append_child(p, gap);
        let b = doc.create_element_typed("strong", "");
        doc.append_child(p, b);
        doc.set_text_content(b, "two");

        let runs = doc.inline_runs(p);
        assert_eq!(runs.len(), 3);
        assert_eq!(runs[1].text, " ", "the gap survives as a single space");
    }

    #[test]
    fn setting_text_content_replaces_the_children_with_one_text_node() {
        // The spec's own definition, and what makes `p.textContent = "x"` and
        // building the same tree by hand agree.
        let mut doc = Document::new("t");
        let p = doc.create_element_typed("p", "");
        doc.append_child(DOCUMENT, p);
        let strong = doc.create_element_typed("strong", "");
        doc.append_child(p, strong);
        doc.set_text_content(strong, "B");
        assert_eq!(doc.inline_runs(p).len(), 1);

        doc.set_text_content(p, "plain");
        assert_eq!(doc.text_content(p), "plain");
        let runs = doc.inline_runs(p);
        assert_eq!(runs.len(), 1, "the <strong> is gone, not appended to");
        assert_eq!(runs[0].text, "plain");
        assert_eq!(runs[0].font.weight, 400);
    }

    #[test]
    fn a_text_node_takes_its_parents_style_because_it_declares_none() {
        // Not a shortcut — the definition. `color` on a `<p>` colours the
        // paragraph's text because the text node inherits, and a text node
        // inherits everything since it declares nothing.
        let mut doc = Document::new("t");
        let p = doc.create_element_typed("p", "");
        doc.append_child(DOCUMENT, p);
        let data = doc.create_text_node("hello");
        doc.append_child(p, data);
        doc.set_style_property(p, "font-size", "27px");

        assert_eq!(doc.inline_runs(p)[0].font.size, 27.0);
    }

    #[test]
    fn editing_a_text_nodes_data_re_derives_the_line() {
        let mut doc = Document::new("t");
        let p = doc.create_element_typed("p", "");
        doc.append_child(DOCUMENT, p);
        let data = doc.create_text_node("before");
        doc.append_child(p, data);
        assert_eq!(doc.inline_runs(p)[0].text, "before");

        doc.set_text_data(data, "after");
        assert_eq!(doc.inline_runs(p)[0].text, "after");
        assert_eq!(doc.text_content(p), "after");
    }

    #[test]
    fn an_inline_child_is_a_run_of_its_parents_text_not_a_box() {
        // The inline formatting context. `<strong>` has no rect, no position
        // and no widget in the tree — it is a differently-styled stretch of the
        // paragraph's line, which is why asking a toolkit for "the strong
        // widget" is the wrong question.
        let mut doc = Document::new("t");
        let p = doc.create_element_typed("p", "");
        doc.append_child(DOCUMENT, p);
        doc.set_text_content(p, "a ");
        let strong = doc.create_element_typed("strong", "");
        doc.append_child(p, strong);
        doc.set_text_content(strong, "b");

        assert!(
            doc.is_inline_content(strong),
            "an inline element is a run of its parent's line, not a box"
        );
        // …and yet it is in the tree, styleable and readable.
        assert_eq!(doc.text_content(strong), "b");
        assert_eq!(doc.style_properties(strong).font_weight, Some(700));

        // TWO runs, not one: the paragraph's own "a " is a text NODE now, so it
        // is a run in its own right rather than a string beside the list.
        let runs = doc.inline_runs(p);
        assert_eq!(runs.len(), 2);
        assert_eq!(runs[0].text, "a ");
        assert_eq!(runs[1].text, "b");
        assert_eq!(
            runs[1].font.weight, 700,
            "the run carries the RESOLVED style — the cascade already ran"
        );
    }

    #[test]
    fn taking_an_inline_element_out_of_flow_makes_it_a_box() {
        // CSS 2.1 §9.7 blockification. There is no line for an absolutely
        // positioned `<span>` to be a run of, so it becomes a box with a rect
        // — which is also why the z-index tests can still use spans as the
        // only widget that paints a background.
        let mut doc = Document::new("t");
        let p = doc.create_element_typed("p", "");
        doc.append_child(DOCUMENT, p);
        let span = doc.create_element_typed("span", "");
        doc.append_child(p, span);
        assert!(doc.is_inline_content(span), "in flow, it is a run");

        doc.set_style_property(span, "position", "absolute");
        assert!(
            !doc.is_inline_content(span),
            "out of flow, it is a box"
        );
    }

    #[test]
    fn restyling_an_inline_child_re_derives_the_line_it_belongs_to() {
        // A run holds a copy of its style by necessity: shaping needs the whole
        // line at once. So the copy has to be re-derived whenever the cascade
        // moves, or the paragraph paints yesterday's answer.
        let mut doc = Document::new("t");
        let p = doc.create_element_typed("p", "");
        doc.append_child(DOCUMENT, p);
        let span = doc.create_element_typed("span", "");
        doc.append_child(p, span);
        doc.set_text_content(span, "x");
        assert_eq!(doc.inline_runs(p)[0].font.size, 14.0);

        // Declared on the PARENT, so it reaches the run only by inheritance.
        doc.set_style_property(p, "font-size", "30px");
        assert_eq!(
            doc.inline_runs(p)[0].font.size,
            30.0,
            "the run takes what the inline element inherits"
        );
    }

    #[test]
    fn removing_the_last_inline_child_empties_the_line() {
        // The stale-copy failure this migration exists to remove: a box that
        // kept painting the runs it was last told.
        let mut doc = Document::new("t");
        let p = doc.create_element_typed("p", "");
        doc.append_child(DOCUMENT, p);
        let span = doc.create_element_typed("span", "");
        doc.append_child(p, span);
        doc.set_text_content(span, "x");
        assert_eq!(doc.inline_runs(p).len(), 1);

        doc.remove_child(p, span);
        assert!(
            doc.inline_runs(p).is_empty(),
            "the run must go with the element"
        );
    }

    /// Put a stylesheet in the document, the way a page does.
    fn with_stylesheet(doc: &mut Document, css: &str) {
        let style = doc.create_element_typed("style", "");
        doc.append_child(DOCUMENT, style);
        doc.set_text_content(style, css);
    }

    #[test]
    fn a_canvas_sizes_its_bitmap_from_the_content_attributes() {
        // HTML §4.12.5. There was no `width`/`height` arm at all, so
        // `<canvas width="640">` stored an inert attribute and the surface kept
        // its default — which is the first thing SDL and `CreateGraphics` set.
        let mut doc = Document::new("t");
        let canvas = doc.create_element_typed("canvas", "");
        doc.append_child(DOCUMENT, canvas);
        let rect = doc.get_bounding_client_rect(canvas).expect("in the document");
        assert_eq!(
            (rect.w, rect.h),
            (300.0, 150.0),
            "the spec's default, not the panel group's 200x150"
        );

        doc.set_attribute(canvas, "width", "640");
        doc.set_attribute(canvas, "height", "480");
        let rect = doc.get_bounding_client_rect(canvas).expect("in the document");
        assert_eq!((rect.w, rect.h), (640.0, 480.0));
        assert_eq!(
            doc.get_attribute(canvas, "width").as_deref(),
            Some("640"),
            "and it reads back as the content attribute it is"
        );
    }

    #[test]
    fn sizing_a_canvas_clears_it_even_to_the_same_value() {
        // `canvas.width = canvas.width` is the idiomatic full reset, and it
        // only works because the spec clears the bitmap on ANY set — not on a
        // change.
        let mut doc = Document::new("t");
        let canvas = doc.create_element_typed("canvas", "");
        doc.append_child(DOCUMENT, canvas);
        doc.set_attribute(canvas, "width", "200");

        use crate::canvas::Canvas as _;
        let surface = doc.canvas_mut(canvas).expect("a canvas widget");
        surface.canvas_mut().fill_rect(0.0, 0.0, 10.0, 10.0);
        assert!(!doc.canvas_mut(canvas).expect("still there").canvas_mut().is_empty());

        doc.set_attribute(canvas, "width", "200");
        assert!(
            doc.canvas_mut(canvas)
                .expect("still there")
                .canvas_mut()
                .is_empty(),
            "setting the same width still clears the bitmap"
        );
    }

    #[test]
    fn width_on_a_non_canvas_is_a_presentational_hint_not_a_bitmap() {
        // The same attributes mean something else on `<img>` and friends: a
        // hint that maps to CSS, sizing the BOX.
        let mut doc = Document::new("t");
        let img = doc.create_element_typed("img", "");
        doc.append_child(DOCUMENT, img);
        doc.set_attribute(img, "width", "120");
        assert_eq!(
            doc.style_properties(img).width,
            Some(crate::css::Length::Px(120.0))
        );
    }

    #[test]
    fn a_canvas_element_hands_out_the_surface_that_paints_it() {
        // The hop between two trees. Finding the node was never the gap —
        // reaching the widget from it was, so paint ops went to a canvas in
        // another tree and nothing rendered.
        let mut doc = Document::new("t");
        let canvas = doc.create_element_typed("canvas", "");
        doc.set_attribute(canvas, "id", "surface");
        doc.append_child(DOCUMENT, canvas);

        assert!(
            doc.canvas_mut(canvas).is_some(),
            "a <canvas> element IS a canvas widget"
        );
        let label = doc.create_element_typed("span", "");
        doc.append_child(DOCUMENT, label);
        assert!(
            doc.canvas_mut(label).is_none(),
            "and nothing else pretends to be one"
        );
    }

    #[test]
    fn a_control_name_resolves_by_name_then_id_then_widget_name() {
        // `getContext("x")` is handed a NAME — SDL, .NET `CreateGraphics` and
        // Flutter all pass one, and none of them set an id. All three
        // identities an element can be known by are answered, in the order a
        // caller means them.
        let mut doc = Document::new("t");
        let named = doc.create_element_typed("canvas", "");
        doc.set_attribute(named, "name", "board");
        doc.append_child(DOCUMENT, named);
        let identified = doc.create_element_typed("canvas", "");
        doc.set_attribute(identified, "id", "surface");
        doc.append_child(DOCUMENT, identified);

        assert_eq!(doc.element_by_control_name("board"), Some(named));
        assert_eq!(doc.element_by_control_name("surface"), Some(identified));
        // The host bridge lower-cases a target on the way in.
        assert_eq!(doc.element_by_control_name("BOARD"), Some(named));
        // And a handle that round-tripped through the toolkit carries the
        // internal widget name.
        assert_eq!(
            doc.element_by_control_name(&Document::widget_name(named)),
            Some(named)
        );
        assert_eq!(doc.element_by_control_name("nothing"), None);
        assert_eq!(doc.element_by_control_name(""), None);
    }

    #[test]
    fn query_selector_answers_in_tree_order_over_the_live_document() {
        // `querySelectorAll` used to be tag selectors only, because the engine
        // that could answer more lived above the seam, wired to a DOM that
        // renders nothing.
        let mut doc = Document::new("t");
        let panel = container_with(&mut doc, 300.0, 200.0);
        let first = doc.create_element_typed("p", "");
        doc.set_attribute(first, "class", "lead");
        doc.append_child(panel, first);
        let second = doc.create_element_typed("p", "");
        doc.set_attribute(second, "class", "lead");
        doc.append_child(panel, second);
        let outside = doc.create_element_typed("p", "");
        doc.set_attribute(outside, "class", "lead");
        doc.append_child(DOCUMENT, outside);

        assert_eq!(doc.query_selector_all(".lead").len(), 3);
        assert_eq!(
            doc.query_selector_all("div > .lead"),
            vec![first, second],
            "the combinator excludes the one outside the panel"
        );
        assert_eq!(
            doc.query_selector("div .lead"),
            Some(first),
            "first in TREE order, not first matched"
        );
        // A selector list is a union.
        assert_eq!(doc.query_selector_all("div, .lead").len(), 4);
    }

    #[test]
    fn an_unsupported_selector_matches_nothing_rather_than_everything() {
        // The spec throws SyntaxError. Refusing to match is the closest answer
        // without an exception channel, and it fails in the safe direction: a
        // call site expecting a few elements must not receive all of them.
        let mut doc = Document::new("t");
        let a = doc.create_element_typed("a", "");
        doc.append_child(DOCUMENT, a);
        assert!(doc.query_selector_all("a:hover").is_empty());
        assert!(doc.query_selector_all("((").is_empty());
        assert_eq!(doc.query_selector_all("a").len(), 1);
    }

    #[test]
    fn a_style_element_is_a_third_cascade_origin() {
        // Until selectors existed the cascade was UA and inline `style=""` with
        // nothing between them — not a missing feature so much as a missing
        // input, since a rule needs a selector.
        let mut doc = Document::new("t");
        with_stylesheet(&mut doc, "p { color: #ff0000; font-size: 30px }");
        let p = doc.create_element_typed("p", "");
        doc.append_child(DOCUMENT, p);

        assert_eq!(doc.style_properties(p).font_size, Some(30.0));
        assert!(doc.style_properties(p).color.is_some());
        // …and it is NOT an inline style: `element.style` must stay empty or a
        // serialiser would write the rule onto every element it matched.
        assert_eq!(doc.get_style_property(p, "font-size"), "");
    }

    #[test]
    fn the_origin_order_is_ua_then_rules_then_inline() {
        // Three layers, each beating the one before it, asserted on ONE
        // property so the order is unambiguous.
        let mut doc = Document::new("t");
        with_stylesheet(&mut doc, "strong { font-weight: 600 }");
        let s = doc.create_element_typed("strong", "");
        doc.append_child(DOCUMENT, s);
        assert_eq!(
            doc.style_properties(s).font_weight,
            Some(600),
            "an author rule beats the UA sheet's bold"
        );

        doc.set_style_property(s, "font-weight", "300");
        assert_eq!(
            doc.style_properties(s).font_weight,
            Some(300),
            "and an inline declaration beats the rule"
        );
    }

    #[test]
    fn specificity_decides_and_source_order_breaks_the_tie() {
        // The whole reason a stylesheet needs more than document order.
        let mut doc = Document::new("t");
        with_stylesheet(
            &mut doc,
            "#target { font-size: 40px } p { font-size: 10px } p { font-size: 20px }",
        );
        let p = doc.create_element_typed("p", "");
        doc.set_attribute(p, "id", "target");
        doc.append_child(DOCUMENT, p);
        assert_eq!(
            doc.style_properties(p).font_size,
            Some(40.0),
            "#id wins however it is ordered"
        );

        let plain = doc.create_element_typed("p", "");
        doc.append_child(DOCUMENT, plain);
        assert_eq!(
            doc.style_properties(plain).font_size,
            Some(20.0),
            "equal specificity: the LAST rule wins"
        );
    }

    #[test]
    fn a_rule_selects_by_class_and_by_relationship() {
        let mut doc = Document::new("t");
        with_stylesheet(&mut doc, "div .lead { font-size: 25px }");
        let panel = container_with(&mut doc, 300.0, 200.0);
        let lead = doc.create_element_typed("p", "");
        doc.set_attribute(lead, "class", "lead intro");
        doc.append_child(panel, lead);

        let orphan = doc.create_element_typed("p", "");
        doc.set_attribute(orphan, "class", "lead");
        doc.append_child(DOCUMENT, orphan);

        assert_eq!(doc.style_properties(lead).font_size, Some(25.0));
        assert_eq!(
            doc.style_properties(orphan).font_size,
            None,
            "the same class outside a div is not selected"
        );
    }

    #[test]
    fn an_unsupported_rule_is_dropped_whole_rather_than_applied_broadly() {
        // CSS error handling. A `:hover` rule whose pseudo-class was silently
        // skipped would apply to every `<a>` unconditionally — worse than never
        // applying, and invisible.
        let mut doc = Document::new("t");
        with_stylesheet(
            &mut doc,
            "a:hover { font-size: 50px } a { font-size: 12px }",
        );
        let a = doc.create_element_typed("a", "");
        doc.append_child(DOCUMENT, a);
        assert_eq!(doc.style_properties(a).font_size, Some(12.0));
    }

    #[test]
    fn comments_and_at_rules_do_not_leak_declarations() {
        // An `@media` whose condition was never evaluated must not apply its
        // contents — skipping the rule is the compliant answer, applying the
        // inside is the tempting wrong one.
        let mut doc = Document::new("t");
        with_stylesheet(
            &mut doc,
            "/* p { font-size: 99px } */ @media print { p { font-size: 88px } } p { font-size: 14px }",
        );
        let p = doc.create_element_typed("p", "");
        doc.append_child(DOCUMENT, p);
        assert_eq!(doc.style_properties(p).font_size, Some(14.0));
    }

    #[test]
    fn an_em_is_relative_to_the_font_it_actually_inherits() {
        // The bug inheritance exposed. Before there was a computed `font-size`
        // to be relative TO, `em` resolved against a hardcoded 16px — so the UA
        // sheet's `h1 { font-size: 2em }` meant 32px inside a 10px container
        // and inside a 40px one alike.
        let mut doc = Document::new("t");
        let small = container_with(&mut doc, 400.0, 300.0);
        doc.set_style_property(small, "font-size", "10px");
        let h_small = doc.create_element_typed("h1", "");
        doc.append_child(small, h_small);

        let big = container_with(&mut doc, 400.0, 300.0);
        doc.set_style_property(big, "font-size", "40px");
        let h_big = doc.create_element_typed("h1", "");
        doc.append_child(big, h_big);

        assert_eq!(doc.style_properties(h_small).font_size, Some(20.0));
        assert_eq!(doc.style_properties(h_big).font_size, Some(80.0));
    }

    #[test]
    fn a_margin_in_em_is_relative_to_the_elements_own_size_not_its_parents() {
        // The subtlety that makes the cascade two passes. `h1 { font-size: 2em;
        // margin: 0.67em 0 }` — the SIZE is twice the container's text, and the
        // margin is two thirds of the heading's own resolved size, not of what
        // it inherited. One pass gets the size right and the margin wrong.
        let mut doc = Document::new("t");
        let panel = container_with(&mut doc, 400.0, 300.0);
        doc.set_style_property(panel, "font-size", "10px");
        let h = doc.create_element_typed("h1", "");
        doc.append_child(panel, h);

        assert_eq!(doc.style_properties(h).font_size, Some(20.0));
        assert_eq!(
            doc.box_edges(h).margin.top,
            0.67 * 20.0,
            "0.67em of the heading's OWN 20px, not of the inherited 10px"
        );
    }

    #[test]
    fn the_widget_and_the_store_resolve_a_font_relative_length_the_same_way() {
        // Two parsers, one question — the shape the colour parser had. The
        // widget path hardcoded `em` to 16px and read `pt` as 4/3 rather than
        // 96/72, so a declaration's USED value depended on which of two
        // functions it reached.
        let mut doc = Document::new("t");
        let panel = container_with(&mut doc, 400.0, 300.0);
        doc.set_style_property(panel, "font-size", "20px");
        let b = doc.create_element_typed("button", "");
        doc.append_child(panel, b);
        doc.set_style_property(b, "width", "3em");

        // The store's answer…
        assert_eq!(
            doc.style_properties(b).width,
            Some(crate::css::Length::Px(60.0))
        );
        // …and the widget's, which is the rect it actually occupies.
        assert_eq!(
            doc.get_bounding_client_rect(b).expect("in the document").w,
            60.0,
            "the control is as wide as the cascade says, not 3 x 16"
        );
    }

    #[test]
    fn a_rem_ignores_every_ancestor_but_the_root() {
        // The whole reason the unit exists, and indistinguishable from `em`
        // while both were 16px constants.
        let mut doc = Document::new("t");
        doc.set_style_property(DOCUMENT, "font-size", "10px");
        let panel = container_with(&mut doc, 400.0, 300.0);
        doc.set_style_property(panel, "font-size", "40px");
        let child = doc.create_element_typed("button", "");
        doc.append_child(panel, child);

        doc.set_style_property(child, "width", "2rem");
        assert_eq!(
            doc.style_properties(child).width,
            Some(crate::css::Length::Px(20.0)),
            "2rem is twice the ROOT's 10px, whatever the 40px parent says"
        );
        doc.set_style_property(child, "width", "2em");
        assert_eq!(
            doc.style_properties(child).width,
            Some(crate::css::Length::Px(80.0)),
            "2em is twice the inherited 40px"
        );
    }

    #[test]
    fn a_heading_is_a_line_of_text_not_a_layout_region() {
        // Same widget as a `<div>`, nothing like the same starting geometry.
        let mut doc = Document::new("t");
        let h = doc.create_element_typed("h1", "");
        doc.append_child(DOCUMENT, h);
        let div = doc.create_element_typed("div", "");
        doc.append_child(DOCUMENT, div);

        let h_rect = doc.get_bounding_client_rect(h).expect("in the document");
        let div_rect = doc.get_bounding_client_rect(div).expect("in the document");
        assert!(
            h_rect.h < div_rect.h,
            "a heading starts one line high ({}), a layout box a region ({})",
            h_rect.h,
            div_rect.h
        );
    }

    #[test]
    fn an_alignment_reaches_the_control_and_is_inherited_like_every_other_text_property() {
        // The calculator's display declares `Alignment := taRightJustify` and
        // its text sat on the left, because `alignment` was an unmapped role
        // and wrote an attribute no element reads. Both halves are asserted:
        // the control the declaration names, and a control that declares
        // nothing under a form that does.
        let mut doc = Document::new("t");
        let form = container_with(&mut doc, 300.0, 200.0);
        let display = doc.create_element_typed("input", "text");
        doc.append_child(form, display);
        doc.set_style_property(display, "text-align", "right");

        assert_eq!(
            doc.style_properties(display).text_align,
            Some(crate::css::TextAlign::Right)
        );

        let label = doc.create_element_typed("label", "");
        doc.append_child(form, label);
        doc.set_style_property(form, "text-align", "center");
        assert_eq!(
            doc.style_properties(label).text_align,
            Some(crate::css::TextAlign::Center),
            "a label that declared nothing takes the form's alignment"
        );
        assert_eq!(
            doc.style_properties(display).text_align,
            Some(crate::css::TextAlign::Right),
            "and the display keeps its own"
        );
    }

    #[test]
    fn a_font_declared_once_on_a_container_reaches_a_control_three_levels_down() {
        // VCL's `ParentFont`/`ParentColor` IS this, and it is the reason a
        // Delphi form declares its font once instead of on every control. Each
        // element here declares NOTHING — the whole value comes from above,
        // which is what separates inheritance from a stamped-on default.
        let mut doc = Document::new("t");
        let form = container_with(&mut doc, 400.0, 300.0);
        doc.set_style_property(form, "color", "#c00000");
        doc.set_style_property(form, "font-family", "Georgia");
        doc.set_style_property(form, "font-size", "18px");

        let outer = container_with(&mut doc, 300.0, 200.0);
        doc.append_child(form, outer);
        let inner = container_with(&mut doc, 200.0, 100.0);
        doc.append_child(outer, inner);
        let button = doc.create_element_typed("button", "");
        doc.append_child(inner, button);

        let props = doc.style_properties(button);
        assert_eq!(props.font_family.as_deref(), Some("Georgia"));
        assert_eq!(props.font_size, Some(18.0));
        assert_eq!(
            props.color,
            doc.style_properties(form).color,
            "the colour is the form's, not a default the button chose"
        );
        assert_eq!(
            doc.get_style_property(button, "font-family"),
            "",
            "and none of it is an inline style — the button declared nothing"
        );
    }

    #[test]
    fn a_user_agent_rule_outranks_what_the_parent_passed_down() {
        // The half that a `div`-to-`span` test cannot see. Inheritance is the
        // FLOOR of the cascade: a `<strong>` inside a `font-weight: 400`
        // container is still bold, because the UA sheet declares and the
        // parent only inherits. Get this backwards and every UA rule in the
        // table stops working the moment anything above it declares the same
        // property — which is every form in the corpus.
        let mut doc = Document::new("t");
        let form = container_with(&mut doc, 200.0, 100.0);
        doc.set_style_property(form, "font-weight", "400");
        doc.set_style_property(form, "color", "#c00000");

        let strong = doc.create_element_typed("strong", "");
        doc.append_child(form, strong);
        let plain = doc.create_element_typed("span", "");
        doc.append_child(form, plain);

        assert_eq!(
            doc.style_properties(strong).font_weight,
            Some(700),
            "the UA rule wins over the inherited weight"
        );
        assert_eq!(
            doc.style_properties(plain).font_weight,
            Some(400),
            "a tag with no rule of its own takes the parent's"
        );
        // `color` has no UA rule on `strong`, so the SAME element inherits one
        // property and overrides the other. That is per-property, not
        // per-element, and it is what makes a cascade a cascade.
        assert_eq!(
            doc.style_properties(strong).color,
            doc.style_properties(form).color
        );
    }

    #[test]
    fn declaring_a_colour_on_the_parent_afterwards_still_reaches_the_child() {
        // The pull side alone is not enough: the widget was told its colour
        // when the declaration was written, and a write on an ancestor happens
        // long after the descendants exist. A form's font is set in
        // `FormCreate`, after every control is constructed.
        let mut doc = Document::new("t");
        let form = container_with(&mut doc, 200.0, 100.0);
        let label = doc.create_element_typed("span", "");
        doc.append_child(form, label);
        assert_eq!(doc.style_properties(label).font_size, None);

        doc.set_style_property(form, "font-size", "22px");

        assert_eq!(
            doc.style_properties(label).font_size,
            Some(22.0),
            "a later write on the parent reaches a child that already existed"
        );
    }

    #[test]
    fn an_elements_own_declaration_is_not_overwritten_by_what_it_inherits() {
        // The ordering hazard inside `recompute_subtree`: it re-pushes the
        // inherited value and the element's own `var()` declarations on the
        // same pass, so the wrong order would let a parent's font size
        // silently replace one the child declared for itself.
        let mut doc = Document::new("t");
        let form = container_with(&mut doc, 200.0, 100.0);
        let label = doc.create_element_typed("span", "");
        doc.append_child(form, label);
        doc.set_style_property(label, "font-size", "9px");

        doc.set_style_property(form, "font-size", "40px");

        assert_eq!(
            doc.style_properties(label).font_size,
            Some(9.0),
            "the child declared 9px and keeps it"
        );
    }

    #[test]
    fn asymmetric_padding_indents_only_the_sides_it_names() {
        // `padding: 0 40px` — a `<ul>`'s marker indent, and the exact case the
        // old code could not express: it flattened the shorthand to its LARGEST
        // edge, so this became a 40px inset on all four sides and the first
        // child dropped 40px down the panel.
        let mut doc = Document::new("t");
        let panel = container_with(&mut doc, 200.0, 200.0);
        let child = doc.create_element_typed("button", "");
        doc.append_child(panel, child);
        doc.set_style_property(panel, "padding", "0 40px");

        let panel_rect = doc.get_bounding_client_rect(panel).expect("the panel is in the document");
        let rect = doc.get_bounding_client_rect(child).expect("the child is in the document");
        assert_eq!(
            rect.x - panel_rect.x,
            40.0,
            "the horizontal padding indents the child"
        );
        assert_eq!(
            rect.y - panel_rect.y,
            0.0,
            "the vertical padding is zero, so the child starts at the top edge"
        );
    }

    #[test]
    fn a_fieldsets_frame_is_real_geometry_not_just_paint() {
        // The first UA rule with a non-zero border-width, and therefore the
        // first time the padding box separates from the border box in real
        // markup: a child at `left: 0` starts INSIDE the frame rather than on
        // top of it. Both CSS and VCL agree — a `TGroupBox`'s client area
        // excludes its bevel.
        let mut doc = Document::new("t");
        let fieldset = doc.create_element_typed("fieldset", "");
        doc.append_child(DOCUMENT, fieldset);
        doc.set_style_property(fieldset, "position", "absolute");
        doc.set_style_property(fieldset, "left", "50px");
        doc.set_style_property(fieldset, "top", "30px");

        assert_eq!(
            doc.box_edges(fieldset).border.left,
            1.0,
            "the UA rule has to reach the element for any of this to matter"
        );

        let child = doc.create_element_typed("button", "");
        doc.append_child(fieldset, child);
        doc.set_style_property(child, "position", "absolute");
        doc.set_style_property(child, "left", "0px");
        doc.set_style_property(child, "top", "0px");

        let rect = doc.get_bounding_client_rect(child).expect("the child is in the document");
        assert_eq!(
            (rect.x, rect.y),
            (51.0, 31.0),
            "one pixel inside the frame on both axes"
        );
    }

    #[test]
    fn box_sizing_decides_what_a_declared_width_measures() {
        // The default is CSS's: a `<div>` with no declaration is content-box.
        // `border-box` is the case that has to be asked for — by an author, or
        // by the UA control rule (see `a_control_measures_its_outer_box`).
        let mut doc = Document::new("t");
        let panel = container_with(&mut doc, 400.0, 300.0);

        let border_box = doc.create_element_typed("div", "");
        doc.append_child(panel, border_box);
        doc.set_style_property(border_box, "position", "absolute");
        doc.set_style_property(border_box, "box-sizing", "border-box");
        doc.set_style_property(border_box, "padding", "10px");
        doc.set_style_property(border_box, "border-width", "2px");
        doc.set_style_property(border_box, "width", "100px");

        // No `box-sizing` declaration at all — the initial value is enough.
        let content_box = doc.create_element_typed("div", "");
        doc.append_child(panel, content_box);
        doc.set_style_property(content_box, "position", "absolute");
        doc.set_style_property(content_box, "padding", "10px");
        doc.set_style_property(content_box, "border-width", "2px");
        doc.set_style_property(content_box, "width", "100px");

        assert_eq!(
            doc.get_bounding_client_rect(border_box).expect("in the document").w,
            100.0,
            "border-box: the 100 INCLUDES the padding and border"
        );
        assert_eq!(
            doc.get_bounding_client_rect(content_box).expect("in the document").w,
            124.0,
            "content-box: the 100 is the content, plus 2×10 padding and 2×2 border"
        );
        assert_eq!(
            doc.content_rect(content_box).w,
            100.0,
            "and the content box is the 100 that was asked for"
        );
    }

    #[test]
    fn a_control_measures_its_outer_box_without_declaring_anything() {
        // The toolkit convention survives the CSS-shaped default because it is
        // a UA declaration on the controls, not an inverted initial value.
        // `TEdit.Width = 100` is a 100px control either way.
        let mut doc = Document::new("t");
        let panel = container_with(&mut doc, 400.0, 300.0);

        let field = doc.create_element_typed("input", "text");
        doc.append_child(panel, field);
        doc.set_style_property(field, "position", "absolute");
        doc.set_style_property(field, "padding", "10px");
        doc.set_style_property(field, "width", "100px");

        assert_eq!(
            doc.get_bounding_client_rect(field).expect("in the document").w,
            100.0,
            "a control's declared width is its OUTER width, from the UA rule"
        );
        assert_eq!(
            doc.content_rect(field).w,
            80.0,
            "and its interior is what the padding leaves"
        );
        // The rule is a UA declaration, so it is NOT an inline style.
        assert_eq!(
            doc.style(field).map(|s| s.css_text()).unwrap_or_default(),
            "padding: 10px; position: absolute; width: 100px",
            "the UA rule must not leak into element.style"
        );
    }

    #[test]
    fn the_four_boxes_coincide_until_an_edge_is_declared() {
        // The model this replaced, stated as a property rather than assumed:
        // one rect per element is CORRECT while every edge is zero, which is
        // why the whole corpus renders identically after the split.
        let mut doc = Document::new("t");
        let panel = container_with(&mut doc, 200.0, 100.0);

        let border = doc.border_rect(panel);
        for (name, rect) in [
            ("padding", doc.padding_rect(panel)),
            ("content", doc.content_rect(panel)),
            ("margin", doc.margin_rect(panel)),
        ] {
            assert_eq!(
                (rect.x, rect.y, rect.w, rect.h),
                (border.x, border.y, border.w, border.h),
                "the {name} box must coincide with the border box while no edge \
                 is declared"
            );
        }
    }

    #[test]
    fn each_edge_moves_exactly_one_boundary() {
        // Four rects are only worth having if they answer DIFFERENTLY. Each
        // edge steps from one box to the next and touches nothing outside it:
        // border separates border from padding, padding separates padding from
        // content, margin grows outward and leaves the border box alone.
        let mut doc = Document::new("t");
        let panel = container_with(&mut doc, 200.0, 100.0);
        doc.set_style_property(panel, "border-width", "2px");
        doc.set_style_property(panel, "padding", "5px");
        doc.set_style_property(panel, "margin", "8px");

        let border = doc.border_rect(panel);
        let padding = doc.padding_rect(panel);
        let content = doc.content_rect(panel);
        let margin = doc.margin_rect(panel);

        assert_eq!(
            (padding.x - border.x, padding.w - border.w),
            (2.0, -4.0),
            "the padding box is the border box inset by the BORDER width"
        );
        assert_eq!(
            (content.x - padding.x, content.w - padding.w),
            (5.0, -10.0),
            "the content box is the padding box inset by the PADDING"
        );
        assert_eq!(
            (margin.x - border.x, margin.w - border.w),
            (-8.0, 16.0),
            "the margin box grows OUTWARD from the border box"
        );
    }

    #[test]
    fn an_absolute_child_resolves_against_the_padding_box_not_the_border_box() {
        // CSS 2.1 §10.1: the containing block is the ancestor's PADDING box. So
        // `left: 0` starts inside the ancestor's border — the one case where
        // border box and padding box give different answers, and the reason the
        // distinction is worth modelling at all.
        let mut doc = Document::new("t");
        let panel = container_with(&mut doc, 400.0, 300.0);
        doc.set_style_property(panel, "left", "40px");
        doc.set_style_property(panel, "top", "20px");
        doc.set_style_property(panel, "border-width", "3px");
        // Padding must NOT move the child: it is inside the containing block,
        // not outside it. This half is what a "content box" mistake would fail.
        doc.set_style_property(panel, "padding", "10px");

        let b = doc.create_element_typed("button", "");
        doc.append_child(panel, b);
        doc.set_style_property(b, "position", "absolute");
        doc.set_style_property(b, "left", "0px");
        doc.set_style_property(b, "top", "0px");

        let rect = doc.get_bounding_client_rect(b).expect("the button is in the document");
        assert_eq!(
            (rect.x, rect.y),
            (43.0, 23.0),
            "the child starts inside the 3px border and is unaffected by the \
             10px padding"
        );
    }

    #[test]
    fn a_percentage_edge_resolves_against_the_width_on_both_axes() {
        // The spec detail that looks like a bug: a percentage padding resolves
        // against the containing block's WIDTH even on the top and bottom sides
        // (CSS 2.1 §8.4). It is what makes `padding: 10%` square, and reading
        // the height for the vertical sides would make it 20 by 10 here.
        let mut doc = Document::new("t");
        let outer = container_with(&mut doc, 200.0, 100.0);
        let inner = doc.create_element_typed("div", "");
        doc.append_child(outer, inner);
        doc.set_style_property(inner, "padding", "10%");

        let edges = doc.box_edges(inner).padding;
        assert_eq!(
            (edges.left, edges.top),
            (20.0, 20.0),
            "10% of the containing block's 200px width, on both axes"
        );
    }

    #[test]
    fn a_declaration_reads_back_as_written_while_layout_gets_a_number() {
        // The CSSOM keeps the declaration and the used value apart, and so must
        // this: `padding: 10%` reads back `10%` from `element.style` and lays
        // out as pixels. Collapsing the two would make the read side lie.
        let mut doc = Document::new("t");
        let outer = container_with(&mut doc, 300.0, 100.0);
        let inner = doc.create_element_typed("div", "");
        doc.append_child(outer, inner);
        doc.set_style_property(inner, "padding", "10%");

        assert_eq!(doc.get_style_property(inner, "padding"), "10%");
        assert_eq!(doc.box_edges(inner).padding.left, 30.0);
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

        let passthrough = doc.create_element_typed("div", "");
        doc.append_child(positioned, passthrough);
        doc.set_style_property(passthrough, "position", "static");
        doc.set_style_property(passthrough, "left", "100px");
        doc.set_style_property(passthrough, "top", "100px");

        let b = doc.create_element_typed("button", "");
        doc.append_child(passthrough, b);
        doc.set_style_property(b, "position", "absolute");
        doc.set_style_property(b, "left", "10px");
        doc.set_style_property(b, "top", "5px");

        let rect = doc.get_bounding_client_rect(b).expect("the button is in the document");
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
        let detached = doc.create_element_typed("input", "text");
        let natural_height = doc.computed_style_property(detached, "height");

        // A flowed sibling, for contrast: it SHOULD stretch.
        let flowed = doc.create_element_typed("input", "text");
        doc.append_child(panel, flowed);
        assert_ne!(doc.computed_style_property(flowed, "height"), natural_height);

        let partly = doc.create_element_typed("input", "text");
        doc.append_child(panel, partly);
        doc.set_style_property(partly, "left", "10px");
        doc.set_style_property(partly, "top", "30px");
        doc.set_style_property(partly, "width", "120px");
        // Width is honoured, height falls back to the control's own — NOT to
        // whatever the flow pass left behind.
        assert_eq!(doc.computed_style_property(partly, "width"), "120px");
        assert_eq!(doc.computed_style_property(partly, "height"), natural_height);
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

        let b = doc.create_element_typed("button", "");
        doc.append_child(panel, b);
        doc.set_style_property(b, "left", "20px");
        doc.set_style_property(b, "top", "50px");
        // 10 + 20, 10 + 50 — not 20, 50 on the body.
        assert_eq!(doc.computed_style_property(b, "left"), "30px");
        assert_eq!(doc.computed_style_property(b, "top"), "60px");
    }

    #[test]
    fn a_static_container_is_transparent_to_its_childrens_coordinates() {
        // The CSS rule, stated as a test so it cannot quietly become a
        // container special case again: only a POSITIONED ancestor is a
        // containing block. A browser handed this markup answers the same, and
        // that equivalence is the whole reason the widget layer is HTML.
        let mut doc = Document::new("t");
        let panel = doc.create_element_typed("div", "");
        doc.append_child(DOCUMENT, panel);
        doc.set_style_property(panel, "position", "static");
        doc.set_style_property(panel, "left", "10px");
        doc.set_style_property(panel, "top", "10px");
        doc.set_style_property(panel, "width", "200px");
        doc.set_style_property(panel, "height", "200px");

        let b = doc.create_element_typed("button", "");
        doc.append_child(panel, b);
        doc.set_style_property(b, "position", "absolute");
        doc.set_style_property(b, "left", "20px");
        // Resolved against the viewport, not the static panel.
        assert_eq!(doc.computed_style_property(b, "left"), "20px");
    }

    #[test]
    fn fixed_resolves_against_the_viewport_not_an_ancestor() {
        let mut doc = Document::new("t");
        let panel = container_with(&mut doc, 200.0, 200.0);
        doc.set_style_property(panel, "left", "10px");
        doc.set_style_property(panel, "top", "10px");
        let b = doc.create_element_typed("button", "");
        doc.append_child(panel, b);
        doc.set_style_property(b, "position", "fixed");
        doc.set_style_property(b, "left", "20px");
        assert_eq!(doc.computed_style_property(b, "left"), "20px");
    }

    #[test]
    fn right_anchors_to_the_opposite_edge() {
        // Flutter's `Positioned` uses all four; VCL and WinForms only ever set
        // two, which is why the other two were never wired.
        let mut doc = Document::new("t");
        let panel = container_with(&mut doc, 400.0, 300.0);
        let b = doc.create_element_typed("button", "");
        doc.append_child(panel, b);
        doc.set_style_property(b, "width", "100px");
        doc.set_style_property(b, "right", "20px");
        // 400 wide, 20 from the right edge, 100 wide → x = 280.
        assert_eq!(doc.computed_style_property(b, "left"), "280px");
    }

    #[test]
    fn left_and_right_together_stretch_the_box_between_them() {
        let mut doc = Document::new("t");
        let panel = container_with(&mut doc, 400.0, 300.0);
        let b = doc.create_element_typed("button", "");
        doc.append_child(panel, b);
        doc.set_style_property(b, "left", "50px");
        doc.set_style_property(b, "right", "50px");
        assert_eq!(doc.computed_style_property(b, "width"), "300px");
    }

    #[test]
    fn constraints_apply_whichever_order_they_arrive_in() {
        // A declarative frontend emits fields in catalog order, so `max-width`
        // may land before or after the width it constrains. Both must clamp.
        let mut doc = Document::new("t");
        let panel = container_with(&mut doc, 400.0, 300.0);

        let after = doc.create_element_typed("button", "");
        doc.append_child(panel, after);
        doc.set_style_property(after, "position", "absolute");
        doc.set_style_property(after, "width", "300px");
        doc.set_style_property(after, "max-width", "120px");
        assert_eq!(doc.computed_style_property(after, "width"), "120px");

        let before = doc.create_element_typed("button", "");
        doc.append_child(panel, before);
        doc.set_style_property(before, "position", "absolute");
        doc.set_style_property(before, "max-width", "120px");
        doc.set_style_property(before, "width", "300px");
        assert_eq!(doc.computed_style_property(before, "width"), "120px");
    }

    #[test]
    fn min_width_wins_over_max_width_when_they_conflict() {
        // CSS resolves the conflict in favour of `min`.
        let mut doc = Document::new("t");
        let panel = container_with(&mut doc, 400.0, 300.0);
        let b = doc.create_element_typed("button", "");
        doc.append_child(panel, b);
        doc.set_style_property(b, "position", "absolute");
        doc.set_style_property(b, "max-width", "50px");
        doc.set_style_property(b, "min-width", "150px");
        doc.set_style_property(b, "width", "100px");
        assert_eq!(doc.computed_style_property(b, "width"), "150px");
    }

    /// **`display: flex` alone is a ROW** — css-flexbox §5.1.
    ///
    /// An undeclared property is not an absent one. Only a declared
    /// `flex-direction` ever reached the widget, whose own default is
    /// `TopDown`, so a container that did not spell the axis out silently got
    /// a column — and every test that could have caught it declared the axis
    /// implicitly by sharing a fixture that meant a column.
    ///
    /// Found in a capture of a `FlowLayoutPanel`, which is `display: flex;
    /// flex-wrap: wrap` and nothing else: it stacked its buttons vertically at
    /// full width where a browser flows them across and wraps.
    #[test]
    fn a_flex_container_that_declares_no_axis_is_a_row() {
        let mut doc = Document::new("t");
        let panel = doc.create_element_typed("div", "");
        doc.append_child(DOCUMENT, panel);
        doc.set_style_property(panel, "display", "flex");
        doc.set_style_property(panel, "width", "400px");
        doc.set_style_property(panel, "height", "200px");

        let a = doc.create_element_typed("button", "");
        let b = doc.create_element_typed("button", "");
        doc.append_child(panel, a);
        doc.append_child(panel, b);

        let first = doc.get_bounding_client_rect(a).unwrap();
        let second = doc.get_bounding_client_rect(b).unwrap();
        assert!(
            second.x > first.x,
            "the second child is BESIDE the first, not below it: {first:?} then {second:?}"
        );
        assert_eq!(
            first.y, second.y,
            "and they share a top edge, which is what one row means"
        );
    }

    #[test]
    fn order_rearranges_children_without_moving_them_in_the_document() {
        let mut doc = Document::new("t");
        let panel = container_with(&mut doc, 200.0, 400.0);
        let first = doc.create_element_typed("button", "");
        let second = doc.create_element_typed("button", "");
        doc.append_child(panel, first);
        doc.append_child(panel, second);
        let first_top = doc.computed_style_property(first, "top");
        // Send the first child to the back.
        doc.set_style_property(first, "order", "5");
        assert_ne!(doc.computed_style_property(first, "top"), first_top);
        assert_eq!(doc.computed_style_property(second, "top"), first_top);
        // …and the document order is untouched.
        assert!(doc.to_html().find("</button>").is_some());
    }

    #[test]
    fn align_self_overrules_the_containers_align_items() {
        let mut doc = Document::new("t");
        let panel = container_with(&mut doc, 200.0, 400.0);
        let a = doc.create_element_typed("button", "");
        let b = doc.create_element_typed("button", "");
        doc.append_child(panel, a);
        doc.append_child(panel, b);
        doc.set_style_property(panel, "align-items", "stretch");
        let stretched = doc.computed_style_property(a, "width");
        doc.set_style_property(b, "align-self", "center");
        assert_eq!(doc.computed_style_property(a, "width"), stretched);
        assert_ne!(doc.computed_style_property(b, "width"), stretched);
    }

    #[test]
    fn justify_content_moves_the_children_within_the_leftover() {
        // Flutter `mainAxisAlignment`. Only observable when nothing grows —
        // with a growing child there is no leftover to distribute.
        let mut doc = Document::new("t");
        let panel = container_with(&mut doc, 200.0, 400.0);
        let a = doc.create_element_typed("button", "");
        doc.append_child(panel, a);
        doc.set_style_property(a, "flex", "0");
        let packed = doc.computed_style_property(a, "top");
        doc.set_style_property(panel, "justify-content", "center");
        let centred = doc.computed_style_property(a, "top");
        assert_ne!(packed, centred, "centring must move the child down");
    }

    #[test]
    fn align_items_stretch_is_the_default_and_stays_so() {
        // Changing the default here would move every existing flutter widget,
        // so this pins it: declaring nothing keeps the old behaviour.
        let mut doc = Document::new("t");
        let panel = container_with(&mut doc, 200.0, 200.0);
        let a = doc.create_element_typed("button", "");
        doc.append_child(panel, a);
        let stretched = doc.computed_style_property(a, "width");
        doc.set_style_property(panel, "align-items", "stretch");
        assert_eq!(doc.computed_style_property(a, "width"), stretched);
        // …and asking for something else visibly differs.
        doc.set_style_property(panel, "align-items", "center");
        assert_ne!(doc.computed_style_property(a, "width"), stretched);
    }

    #[test]
    fn flex_direction_chooses_the_axis() {
        let mut doc = Document::new("t");
        let panel = container_with(&mut doc, 200.0, 200.0);
        let a = doc.create_element_typed("button", "");
        let b = doc.create_element_typed("button", "");
        doc.append_child(panel, a);
        doc.append_child(panel, b);
        doc.set_style_property(panel, "flex-direction", "row");
        // Side by side: same top, different left.
        assert_eq!(doc.computed_style_property(a, "top"), doc.computed_style_property(b, "top"));
        assert_ne!(doc.computed_style_property(a, "left"), doc.computed_style_property(b, "left"));
        // …and the COLUMN direction is what discriminates. Without this the
        // test passed for the wrong reason: normal flow also puts two
        // `inline-block` buttons side by side, so `row` was satisfied whether
        // or not `flex-direction` was read at all. A green test that no longer
        // tells the two apart is worse than a red one — nothing reports it.
        doc.set_style_property(panel, "flex-direction", "column");
        assert_ne!(doc.computed_style_property(a, "top"), doc.computed_style_property(b, "top"));
    }

    #[test]
    fn a_click_arrives_as_a_dom_event() {
        use crate::layout::{MouseButton, MouseEvent, MouseEventKind};
        let mut doc = Document::new("t");
        let b = doc.create_element_typed("button", "");
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

    /// **`:nth-child()` and friends**, including the invalidation — which is
    /// the half a per-node restyle can never reach.
    #[test]
    fn positional_selectors_match_and_survive_an_insertion() {
        let mut doc = Document::new("t");
        let sheet = doc.create_element_typed("style", "");
        let css = doc.create_text_node(
            "p { color: #000000 } p:nth-child(odd) { color: #ff0000 } \
             p:last-child { color: #0000ff }",
        );
        doc.append_child(sheet, css);
        doc.append_child(DOCUMENT, sheet);

        let list = doc.create_element_typed("div", "");
        doc.append_child(DOCUMENT, list);
        let mut rows = Vec::new();
        for _ in 0..3 {
            let row = doc.create_element_typed("p", "");
            doc.append_child(list, row);
            rows.push(row);
        }
        let colour = |doc: &mut Document, n: NodeId| {
            doc.get_computed_style(n).color.map(crate::css::serialize_color)
        };
        assert_eq!(colour(&mut doc, rows[0]), Some("#ff0000".to_string()), "1st is odd");
        assert_eq!(colour(&mut doc, rows[1]), Some("#000000".to_string()), "2nd is even");
        // The third is both odd AND last; `:last-child` is later at equal
        // specificity, so it wins.
        assert_eq!(colour(&mut doc, rows[2]), Some("#0000ff".to_string()), "3rd is last");

        // **Append a fourth.** The third stops being last and becomes plain
        // odd — a change to an element nobody touched.
        let fourth = doc.create_element_typed("p", "");
        doc.append_child(list, fourth);
        assert_eq!(
            colour(&mut doc, rows[2]),
            Some("#ff0000".to_string()),
            "the old last-child kept a style it no longer matches"
        );
        assert_eq!(colour(&mut doc, fourth), Some("#0000ff".to_string()), "the new one is last");
    }

    /// **`white-space: nowrap` keeps text on one line** — in the measurement
    /// as well as the paint.
    ///
    /// The two have to agree: a box measured as if it wrapped is as tall as
    /// text that breaks, and then the renderer draws one line inside it.
    #[test]
    fn nowrap_keeps_text_on_one_line_and_the_height_agrees() {
        let mut doc = Document::new("t");
        let para = doc.create_element_typed("p", "");
        doc.set_style_property(para, "width", "80px");
        doc.append_child(DOCUMENT, para);
        let words = doc.create_text_node(&"wrap ".repeat(40));
        doc.append_child(para, words);

        let wrapped = doc.border_rect(para).h;
        doc.set_style_property(para, "white-space", "nowrap");
        let single = doc.border_rect(para).h;
        assert!(
            single < wrapped,
            "nowrap did not stop the wrapping: {single} vs {wrapped}"
        );

        // …and back, so the property is read rather than latched.
        doc.set_style_property(para, "white-space", "normal");
        assert!(
            (doc.border_rect(para).h - wrapped).abs() < 0.5,
            "the box did not wrap again once allowed to"
        );
    }

    /// **`:hover` styles the element the pointer is over** — and stops when it
    /// leaves.
    ///
    /// Both halves matter. Matching alone is not enough: computed values are
    /// cached, so a rule that matched only when some unrelated pass happened to
    /// re-cascade would apply at random. The state sweep is what re-runs the
    /// cascade for the element that changed, in both directions.
    #[test]
    fn hover_styles_the_element_under_the_pointer_and_releases_it() {
        let mut doc = Document::new("t");
        let sheet = doc.create_element_typed("style", "");
        let css = doc.create_text_node("button { color: #000000 } button:hover { color: #ff0000 }");
        doc.append_child(sheet, css);
        doc.append_child(DOCUMENT, sheet);

        let button = doc.create_element_typed("button", "");
        doc.append_child(DOCUMENT, button);
        doc.refresh_interaction_states();
        assert_eq!(
            doc.get_computed_style(button).color.map(crate::css::serialize_color),
            Some("#000000".to_string()),
            "an unhovered button takes the plain rule"
        );

        // The pointer arrives. Nothing tells the document — the arranging panel
        // sets this on the widget directly, which is why the sweep exists.
        doc.widget_mut(button).expect("the button has a widget").set_hovered(true);
        assert!(
            doc.refresh_interaction_states(),
            "the sweep must report that something moved"
        );
        assert_eq!(
            doc.get_computed_style(button).color.map(crate::css::serialize_color),
            Some("#ff0000".to_string()),
            "`:hover` did not win, though it is more specific and later"
        );

        // …and leaves. A one-way implementation looks correct until you move
        // the mouse away.
        doc.widget_mut(button).expect("still there").set_hovered(false);
        doc.refresh_interaction_states();
        assert_eq!(
            doc.get_computed_style(button).color.map(crate::css::serialize_color),
            Some("#000000".to_string()),
            "the hover style outlived the hover"
        );
    }

    /// **Padding insets a box's contents** — boxes AND text, and a declared
    /// height survives being a flex item.
    ///
    /// Three separate questions that a Flutter `AppBar` asks at once: it
    /// declares `height`, `padding-left` and a `color` its title should
    /// inherit, and it is a `flex: 0` item of the `Scaffold`.
    #[test]
    fn padding_insets_both_boxes_and_text_and_a_declared_height_survives_flex() {
        let mut doc = Document::new("t");
        let bar = doc.create_element_typed("div", "");
        doc.set_style_property(bar, "padding-left", "16px");
        doc.set_style_property(bar, "height", "56px");
        doc.set_style_property(bar, "color", "#ffffff");
        let child = doc.create_element_typed("div", "");
        doc.append_child(bar, child);
        doc.append_child(DOCUMENT, bar);

        assert_eq!(
            doc.border_rect(bar).h,
            56.0,
            "a declared height is the author's answer"
        );
        let (outer, inner) = (doc.border_rect(bar), doc.border_rect(child));
        assert!(
            inner.x >= outer.x + 15.5,
            "padding did not inset the child box: {} vs {}",
            inner.x,
            outer.x
        );

        // **The box's own TEXT is inset by its padding too**, and takes the
        // colour the box declares. An `AppBar`'s title is a child `<span>` that
        // becomes an inline run of the bar, so neither fact reaches it through
        // the child's own box — the run has to carry them.
        let title = doc.create_element_typed("span", "");
        let words = doc.create_text_node("Calculator");
        doc.append_child(title, words);
        doc.append_child(bar, title);
        let runs = doc.inline_runs(bar);
        assert!(!runs.is_empty(), "the bar has no inline content to paint");
        assert_eq!(
            runs[0].color,
            (255, 255, 255, 255),
            "the title did not take the colour its bar declares"
        );

        // …and the same again with the tree built DETACHED and attached last,
        // which is the order every bottom-up frontend uses. Inheritance has to
        // survive it: the run is built when the span joins its parent, and at
        // that moment nothing above is connected.
        let bar2 = doc.create_element_typed("div", "");
        doc.set_style_property(bar2, "color", "#ffffff");
        let title2 = doc.create_element_typed("span", "");
        let words2 = doc.create_text_node("Calculator");
        doc.append_child(title2, words2);
        doc.append_child(bar2, title2);
        doc.append_child(DOCUMENT, bar2);
        let runs2 = doc.inline_runs(bar2);
        assert!(!runs2.is_empty(), "the detached bar has no inline content");
        assert_eq!(
            runs2[0].color,
            (255, 255, 255, 255),
            "a title built while its bar was DETACHED lost the inherited colour"
        );

        // …and once more with the CHILD built FIRST. `AppBar(title: Text('x'))`
        // evaluates its argument before the bar exists, so the span is created,
        // and gets a computed style, while nothing is above it to inherit from.
        // The colour only becomes inheritable when it is later nested.
        let title3 = doc.create_element_typed("span", "");
        let words3 = doc.create_text_node("Calculator");
        doc.append_child(title3, words3);
        // …and on a FLEX bar, which is what an `AppBar` is
        // (`header;display:flex;align-items:center`).
        let bar3 = doc.create_element_typed("header", "");
        doc.set_style_property(bar3, "display", "flex");
        doc.set_style_property(bar3, "align-items", "center");
        doc.set_style_property(bar3, "color", "#ffffff");
        doc.append_child(bar3, title3);
        doc.append_child(DOCUMENT, bar3);
        let runs3 = doc.inline_runs(bar3);
        assert!(!runs3.is_empty(), "the argument-first bar has no content");
        assert_eq!(
            runs3[0].color,
            (255, 255, 255, 255),
            "a title CONSTRUCTED BEFORE ITS PARENT never inherited the colour"
        );

        // The same box as a `flex: 0` item of a flex column — the shape an
        // `AppBar` is in. Its declared height must not be replaced by the
        // container's guess for a non-growing item.
        let column = doc.create_element_typed("div", "");
        doc.set_style_property(column, "display", "flex");
        doc.set_style_property(column, "flex-direction", "column");
        doc.set_style_property(column, "height", "600px");
        let header = doc.create_element_typed("div", "");
        doc.set_style_property(header, "flex", "0");
        doc.set_style_property(header, "height", "56px");
        doc.append_child(column, header);
        doc.append_child(DOCUMENT, column);
        assert_eq!(
            doc.border_rect(header).h,
            56.0,
            "a flex item's DECLARED height was replaced by the container's guess"
        );
    }

    /// **A percentage size is a percentage of the CONTAINING BLOCK** — §10.1,
    /// and for a static box that is its parent's content edge, not the
    /// viewport.
    ///
    /// Declared while the box is still detached, which is how every bottom-up
    /// frontend builds: the value has to survive being resolved before there is
    /// a parent to resolve it against.
    #[test]
    fn a_percentage_size_resolves_against_the_parent_not_the_viewport() {
        let mut doc = Document::new("t");
        let box_ = doc.create_element_typed("div", "");
        doc.set_style_property(box_, "width", "200px");
        doc.set_style_property(box_, "height", "120px");

        let child = doc.create_element_typed("button", "");
        doc.set_style_property(child, "width", "100%");
        doc.set_style_property(child, "height", "100%");
        doc.append_child(box_, child);
        doc.append_child(DOCUMENT, box_);

        // …and again through the shape a Flutter row builds, where the parent's
        // own width is not declared but comes from the flow, and the whole
        // subtree is attached at the root: `Row > Expanded > Padding > button`.
        let row = doc.create_element_typed("div", "");
        doc.set_style_property(row, "display", "flex");
        doc.set_style_property(row, "flex-direction", "row");
        doc.set_style_property(row, "height", "100%");
        let cell = doc.create_element_typed("div", "");
        doc.set_style_property(cell, "flex", "1");
        let pad = doc.create_element_typed("div", "");
        doc.set_style_property(pad, "height", "100%");
        let key = doc.create_element_typed("button", "");
        doc.set_style_property(key, "width", "100%");
        doc.set_style_property(key, "height", "100%");
        doc.append_child(pad, key);
        doc.append_child(cell, pad);
        doc.append_child(row, cell);
        doc.append_child(DOCUMENT, row);
        let pad_box = doc.content_rect(pad);
        let key_box = doc.border_rect(key);
        // ⚠ 8px of slack, and it is a known discrepancy rather than rounding:
        // the layout panels carry a built-in 4px padding that `box_edges` does
        // not report, so the DOM's content box is one panel padding wider than
        // the slot the panel actually gives its child. Closing it means moving
        // that default into the UA sheet, where the box model can see it.
        // Before this test the same measurement was 800 — a whole viewport.
        assert!(
            key_box.w <= pad_box.w + 8.5,
            "the key overflowed its cell: key={key_box:?} padcontent={pad_box:?} \
             cell={:?}",
            doc.border_rect(cell),
        );

        let outer = doc.content_rect(box_);
        let inner = doc.border_rect(child);
        assert!(
            (inner.w - outer.w).abs() < 0.5,
            "width resolved against the wrong box: {} for a {} content box",
            inner.w,
            outer.w
        );
        assert!(
            (inner.h - outer.h).abs() < 0.5,
            "height resolved against the wrong box: {} for a {} content box",
            inner.h,
            outer.h
        );
    }

    /// **A column with `height: auto` grows to hold its rows.**
    ///
    /// The exact shape every Flutter frontend emits — `Column` of `Expanded`
    /// rows of `Expanded` `Padding` of `ElevatedButton` — reduced to the boxes
    /// it becomes. `flex: 1` is `flex: 1 1 0%`, and against a container whose
    /// main size is not known yet that basis resolves to CONTENT, not to zero.
    ///
    /// Get it wrong and the rows divide `default_size`'s 150px guess four ways
    /// — about 37px each, for buttons nearer 28 plus padding — so the rows
    /// print on top of one another at half height. Growing cannot rescue it
    /// either: the content height would be measured FROM the guess, a fixed
    /// point that reproduces 150 for ever.
    #[test]
    fn a_column_of_flex_rows_grows_to_hold_them() {
        let mut doc = Document::new("t");
        let column = doc.create_element_typed("div", "");
        doc.set_style_property(column, "display", "flex");
        doc.set_style_property(column, "flex-direction", "column");

        let mut rows = Vec::new();
        let mut inner_rows = Vec::new();
        let mut buttons = Vec::new();
        for r in 0..4 {
            let cell = doc.create_element_typed("div", "");
            doc.set_style_property(cell, "flex", "1");
            let row = doc.create_element_typed("div", "");
            doc.set_style_property(row, "display", "flex");
            doc.set_style_property(row, "flex-direction", "row");
            for c in 0..4 {
                let item = doc.create_element_typed("div", "");
                doc.set_style_property(item, "flex", "1");
                let pad = doc.create_element_typed("div", "");
                let button = doc.create_element_typed("button", "");
                // A `<span>`, not a text node — Flutter's `Text` widget is an
                // element, and a button's face being a child ELEMENT is the
                // case that differs from a bare caption.
                let span = doc.create_element_typed("span", "");
                let label = doc.create_text_node(&format!("{}", r * 4 + c));
                doc.append_child(span, label);
                doc.append_child(button, span);
                doc.append_child(pad, button);
                doc.append_child(item, pad);
                doc.append_child(row, item);
                if r == 0 && c == 0 {
                    buttons.push(button);
                }
            }
            doc.append_child(cell, row);
            doc.append_child(column, cell);
            rows.push(cell);
            inner_rows.push(row);
        }
        // The wrappers a real Flutter tree puts above the column: a `Scaffold`
        // with an `AppBar`, inside a `MaterialApp`, each a flex column of its
        // own, plus the `Container` whose null width/height the emitter writes
        // out as the uninterpretable `nullpx`.
        let panel = doc.create_element_typed("div", "");
        doc.set_style_property(panel, "height", "nullpx");
        doc.set_style_property(panel, "width", "nullpx");
        let display = doc.create_element_typed("span", "");
        let zero = doc.create_text_node("0");
        doc.append_child(display, zero);
        doc.append_child(panel, display);
        doc.insert_before(column, panel, Some(rows[0]));

        let header = doc.create_element_typed("header", "");
        doc.set_style_property(header, "display", "flex");
        doc.set_style_property(header, "align-items", "center");
        doc.set_style_property(header, "flex", "0");
        let title = doc.create_element_typed("span", "");
        let title_text = doc.create_text_node("Calculator");
        doc.append_child(title, title_text);
        doc.append_child(header, title);

        let scaffold = doc.create_element_typed("div", "");
        doc.set_style_property(scaffold, "display", "flex");
        doc.set_style_property(scaffold, "flex-direction", "column");
        doc.append_child(scaffold, header);
        doc.append_child(scaffold, column);

        let app = doc.create_element_typed("div", "");
        doc.set_style_property(app, "display", "flex");
        doc.set_style_property(app, "flex-direction", "column");
        doc.append_child(app, scaffold);

        // **Attached LAST.** A frontend that builds its tree bottom-up — every
        // Flutter one does — hands the document a finished subtree, so every
        // style on it was declared while it was detached. Attaching first would
        // test a different program from the one that runs.
        doc.append_child(DOCUMENT, app);

        let geometry: Vec<LayoutRect> = rows.iter().map(|&n| doc.border_rect(n)).collect();
        // Every row taller than the button it holds — the failure that drew
        // half of each one.
        for (i, rect) in geometry.iter().enumerate() {
            assert!(
                rect.h >= 24.0,
                "row {i} collapsed to {}: {geometry:?}",
                rect.h
            );
        }
        // And no row starting before the one above it ended.
        for pair in geometry.windows(2) {
            assert!(
                pair[1].y >= pair[0].y + pair[0].h - 0.5,
                "rows overlap: {geometry:?}"
            );
        }
        // **Every box within a padding of what it holds.** Non-overlap alone
        // does not say the sizing is right: a box stuck at `default_size`'s
        // 150px guess overlaps nothing, it just leaves a hole where a
        // calculator's rows should be — which is the shape the overlap turned
        // into once the flex half was fixed. Each of these is one level of the
        // chain, so a break names the level it happened at.
        let button_h = doc.border_rect(buttons[0]).h;
        let inner = doc.border_rect(inner_rows[0]).h;
        assert!(
            inner <= button_h + 24.0,
            "the inner row did not size to its buttons: {inner} for a {button_h} button"
        );
        assert!(
            geometry[0].h <= inner + 12.0,
            "the cell did not size to its row: {} for a {inner} row",
            geometry[0].h
        );
        // The column is as tall as what it flowed, not as tall as its guess —
        // so the last row ends inside it.
        let column_rect = doc.border_rect(column);
        let last = geometry.last().expect("four rows");
        assert!(
            last.y + last.h <= column_rect.y + column_rect.h + 0.5,
            "the column did not grow to its rows: {column_rect:?} vs {last:?}"
        );
    }

    // ── The methods that arrived with the second engine ────────────────────
    //
    // `rhtmledit` carries these same tests under the same names, in
    // `crates/htmlbox/src/tests/test_dom_api.rs`. The two engines are meant to
    // be swappable, and a shared SIGNATURE proves nothing about shared
    // BEHAVIOUR — only a pair of tests that assert the same thing does.
    // The setup differs because this engine has no HTML parser; the assertions
    // must not.

    /// A `<div>` under the document, with `id` set — the closest this engine
    /// gets to writing a line of markup.
    fn element(doc: &mut Document, tag: &str, id: &str) -> NodeId {
        let node = doc.create_element(tag);
        if !id.is_empty() {
            doc.set_id(node, id);
        }
        node
    }

    #[test]
    fn normalize_merges_adjacent_text_and_drops_empties() {
        let mut doc = Document::new("t");
        let d = element(&mut doc, "div", "d");
        doc.append_child(DOCUMENT, d);
        for part in ["a", "", "b"] {
            let t = doc.create_text_node(part);
            doc.append_child(d, t);
        }
        let span = element(&mut doc, "span", "");
        doc.append_child(d, span);
        let tail = doc.create_text_node("c");
        doc.append_child(d, tail);

        doc.normalize(d);

        // "a" and "b" merge THROUGH the empty node between them — a zero-length
        // text node is removed, and removing it does not break adjacency. "c"
        // is separated by an element, so it stays its own node.
        let kinds: Vec<String> = doc
            .child_nodes(d)
            .iter()
            .map(|&c| {
                if doc.node(c).map(|n| n.kind) == Some(NodeKind::Text) {
                    doc.text_data(c)
                } else {
                    format!("<{}>", doc.node_name(c))
                }
            })
            .collect();
        assert_eq!(kinds, vec!["ab", "<span>", "c"]);
    }

    #[test]
    fn appending_a_fragment_moves_its_children_and_not_itself() {
        let mut doc = Document::new("t");
        let d = element(&mut doc, "div", "d");
        doc.append_child(DOCUMENT, d);
        let keep = element(&mut doc, "i", "");
        doc.append_child(d, keep);

        let fragment = doc.create_document_fragment();
        assert!(doc.is_document_fragment(fragment));
        assert_eq!(doc.node_type(fragment), 11);
        assert_eq!(doc.node_name(fragment), "#document-fragment");
        for tag in ["a", "b"] {
            let child = doc.create_element(tag);
            doc.append_child(fragment, child);
        }

        doc.append_child(d, fragment);

        // The fragment is NOT in the tree — its children are, in order, and at
        // the level the caller asked for rather than one deeper.
        let tags: Vec<String> = doc.child_nodes(d).iter().map(|&c| doc.node_name(c)).collect();
        assert_eq!(tags, vec!["i", "a", "b"]);
        assert!(
            !tags.iter().any(|t| t == "#document-fragment"),
            "the fragment itself must not land in the tree"
        );
        // …and it is empty afterwards, having handed its children over.
        assert!(doc.child_nodes(fragment).is_empty());
    }

    #[test]
    fn inserting_a_fragment_splices_it_in_order() {
        let mut doc = Document::new("t");
        let d = element(&mut doc, "div", "d");
        doc.append_child(DOCUMENT, d);
        let pivot = element(&mut doc, "i", "pivot");
        doc.append_child(d, pivot);

        let fragment = doc.create_document_fragment();
        for tag in ["a", "b"] {
            let child = doc.create_element(tag);
            doc.append_child(fragment, child);
        }

        doc.insert_before(d, fragment, Some(pivot));

        // Each child goes before the SAME reference, so they arrive in the
        // order they were in — not reversed.
        let tags: Vec<String> = doc.child_nodes(d).iter().map(|&c| doc.node_name(c)).collect();
        assert_eq!(tags, vec!["a", "b", "i"]);
    }

    #[test]
    fn remove_child_can_remove_a_text_node() {
        // The regression guard for a silent no-op. `removeChild` used to
        // require the child to own a widget, and a text node never does — so
        // it unlinked nothing and reported `false`, and every caller that
        // clears a box by looping over its children left the text behind.
        let mut doc = Document::new("t");
        let p = element(&mut doc, "p", "p");
        doc.append_child(DOCUMENT, p);
        let text = doc.create_text_node("gone");
        doc.append_child(p, text);
        assert_eq!(doc.child_nodes(p).len(), 1);

        assert!(doc.remove_child(p, text), "removing a text node must report success");
        assert!(doc.child_nodes(p).is_empty(), "…and must actually unlink it");

        // Still false for a node that is not a child of this parent — the
        // check that replaced the widget lookup has to keep saying no.
        let stranger = doc.create_element("i");
        assert!(!doc.remove_child(p, stranger));
    }

    #[test]
    fn set_text_content_replaces_rather_than_appends() {
        // The consequence of the bug above, at the level a page would meet it.
        let mut doc = Document::new("t");
        let p = element(&mut doc, "p", "p");
        doc.append_child(DOCUMENT, p);
        doc.set_text_content(p, "first");
        doc.set_text_content(p, "second");
        assert_eq!(doc.text_content(p), "second", "the first text must be gone, not kept beside it");
    }

    #[test]
    fn is_equal_node_ignores_identity_and_attribute_order() {
        let mut doc = Document::new("t");
        let one = doc.create_element("p");
        doc.set_attribute(one, "id", "a");
        doc.set_attribute(one, "class", "x");
        let two = doc.create_element("p");
        // The same two attributes, set in the OPPOSITE order.
        doc.set_attribute(two, "class", "x");
        doc.set_attribute(two, "id", "a");
        for node in [one, two] {
            let t = doc.create_text_node("hi");
            doc.append_child(node, t);
        }

        assert_ne!(one, two, "these must be different NODES");
        assert!(doc.is_equal_node(one, two), "attribute order is not part of equality");

        // Same attributes, different text.
        let three = doc.create_element("p");
        doc.set_attribute(three, "id", "a");
        doc.set_attribute(three, "class", "x");
        let bye = doc.create_text_node("bye");
        doc.append_child(three, bye);
        assert!(!doc.is_equal_node(one, three));

        assert!(doc.is_equal_node(one, one), "a node always equals itself");
    }

    #[test]
    fn compare_document_position_reports_containment_and_order() {
        let mut doc = Document::new("t");
        let outer = element(&mut doc, "div", "outer");
        doc.append_child(DOCUMENT, outer);
        let first = element(&mut doc, "p", "first");
        let second = element(&mut doc, "p", "second");
        doc.append_child(outer, first);
        doc.append_child(outer, second);

        assert_eq!(doc.compare_document_position(outer, outer), 0);
        // CONTAINED_BY | FOLLOWING — containment always carries a direction.
        assert_eq!(doc.compare_document_position(outer, first), 0x10 | 0x04);
        // CONTAINS | PRECEDING, the exact mirror.
        assert_eq!(doc.compare_document_position(first, outer), 0x08 | 0x02);
        assert_eq!(doc.compare_document_position(first, second), 0x04);
        assert_eq!(doc.compare_document_position(second, first), 0x02);
    }

    #[test]
    fn prepend_keeps_the_given_order() {
        let mut doc = Document::new("t");
        let d = element(&mut doc, "div", "d");
        doc.append_child(DOCUMENT, d);
        let z = element(&mut doc, "i", "");
        doc.append_child(d, z);
        let a = doc.create_element("a");
        let b = doc.create_element("b");

        doc.prepend(d, &[a, b]);

        // The naive implementation — insert each before the CURRENT first child
        // — yields b, a, i. The nodes must arrive in the order they were given.
        let tags: Vec<String> = doc.child_nodes(d).iter().map(|&c| doc.node_name(c)).collect();
        assert_eq!(tags, vec!["a", "b", "i"]);
    }

    #[test]
    fn before_after_and_replace_with_place_nodes_around_a_sibling() {
        let mut doc = Document::new("t");
        let d = element(&mut doc, "div", "d");
        doc.append_child(DOCUMENT, d);
        let pivot = element(&mut doc, "i", "pivot");
        doc.append_child(d, pivot);

        let a = doc.create_element("a");
        let b = doc.create_element("b");
        doc.before(pivot, &[a]);
        doc.after(pivot, &[b]);
        let tags: Vec<String> = doc.child_nodes(d).iter().map(|&c| doc.node_name(c)).collect();
        assert_eq!(tags, vec!["a", "i", "b"]);

        let u = doc.create_element("u");
        doc.replace_with(pivot, &[u]);
        let tags: Vec<String> = doc.child_nodes(d).iter().map(|&c| doc.node_name(c)).collect();
        assert_eq!(tags, vec!["a", "u", "b"], "the pivot is gone and `u` sits where it was");
    }

    #[test]
    fn insert_adjacent_element_places_at_all_four_positions() {
        let mut doc = Document::new("t");
        let d = element(&mut doc, "div", "d");
        doc.append_child(DOCUMENT, d);
        let p = element(&mut doc, "p", "p");
        doc.append_child(d, p);
        let x = doc.create_text_node("x");
        doc.append_child(p, x);

        for (position, tag) in [
            ("beforebegin", "a"),
            ("afterbegin", "b"),
            ("beforeend", "u"),
            ("afterend", "s"),
        ] {
            let node = doc.create_element(tag);
            assert_eq!(
                doc.insert_adjacent_element(p, position, node),
                Some(node),
                "{position} should return the inserted element"
            );
        }

        // `beforebegin`/`afterend` are p's SIBLINGS…
        let outer: Vec<String> = doc.child_nodes(d).iter().map(|&c| doc.node_name(c)).collect();
        assert_eq!(outer, vec!["a", "p", "s"]);
        // …and `afterbegin`/`beforeend` are its children, around the text.
        let inner: Vec<String> = doc.child_nodes(p).iter().map(|&c| doc.node_name(c)).collect();
        assert_eq!(inner, vec!["b", "#text", "u"]);

        assert_eq!(doc.insert_adjacent_element(p, "sideways", d), None);
    }

    #[test]
    fn dataset_maps_camel_case_to_data_attributes() {
        let mut doc = Document::new("t");
        let d = element(&mut doc, "div", "d");
        doc.append_child(DOCUMENT, d);
        doc.set_attribute(d, "data-user-name", "ada");
        doc.set_attribute(d, "data-id", "7");
        doc.set_attribute(d, "class", "c");

        // The attribute is `data-user-name`; the IDL name is `userName`.
        assert_eq!(doc.dataset_get(d, "userName").as_deref(), Some("ada"));
        assert_eq!(doc.dataset_get(d, "id").as_deref(), Some("7"));
        // `class` is not `data-*` and must not appear. Asserted against the
        // raw return value — sorting here first would hide an engine that
        // answers in a different order from the other one.
        assert_eq!(
            doc.dataset(d),
            vec![
                ("id".to_string(), "7".to_string()),
                ("userName".to_string(), "ada".to_string()),
            ]
        );

        doc.set_dataset(d, "roleName", "admin");
        assert_eq!(doc.get_attribute(d, "data-role-name").as_deref(), Some("admin"));
        doc.remove_dataset(d, "roleName");
        assert_eq!(doc.dataset_get(d, "roleName"), None);
    }

    #[test]
    fn inner_text_skips_hidden_subtrees() {
        let mut doc = Document::new("t");
        let d = element(&mut doc, "div", "d");
        doc.append_child(DOCUMENT, d);
        for (tag, text, hidden) in [("p", "shown", false), ("p", "hidden", true), ("span", "tail", false)] {
            let child = doc.create_element(tag);
            doc.append_child(d, child);
            if hidden {
                doc.set_style_property(child, "display", "none");
            }
            let t = doc.create_text_node(text);
            doc.append_child(child, t);
        }

        let text = doc.inner_text(d);
        assert!(text.contains("shown"), "got {text:?}");
        assert!(text.contains("tail"), "got {text:?}");
        assert!(
            !text.contains("hidden"),
            "innerText is the RENDERED text — this is the whole difference from textContent: {text:?}"
        );
        // …and textContent still reports everything, unchanged.
        assert!(doc.text_content(d).contains("hidden"));
    }

    #[test]
    fn tab_index_defaults_by_element_kind() {
        let mut doc = Document::new("t");
        let div = element(&mut doc, "div", "");
        let button = element(&mut doc, "button", "");
        let bare = element(&mut doc, "a", "");
        let link = element(&mut doc, "a", "");
        doc.set_attribute(link, "href", "/");
        let explicit = element(&mut doc, "i", "");
        doc.set_attribute(explicit, "tabindex", "3");

        assert_eq!(doc.tab_index(div), -1, "a div is not in the tab sequence by default");
        assert_eq!(doc.tab_index(button), 0, "a button is");
        assert_eq!(doc.tab_index(bare), -1, "an anchor with nowhere to go is not");
        assert_eq!(doc.tab_index(link), 0, "an anchor with an href is");
        assert_eq!(doc.tab_index(explicit), 3, "an explicit tabindex wins over any default");
    }

    #[test]
    fn click_reaches_a_listener_like_a_real_one() {
        use std::sync::atomic::{AtomicUsize, Ordering};
        use std::sync::Arc;

        let mut doc = Document::new("t");
        let button = element(&mut doc, "button", "b");
        doc.append_child(DOCUMENT, button);

        let hits = Arc::new(AtomicUsize::new(0));
        let counter = hits.clone();
        doc.add_event_listener(
            button,
            "click",
            Box::new(move |_event, _doc| {
                counter.fetch_add(1, Ordering::SeqCst);
            }),
        );

        doc.click(button);

        assert_eq!(
            hits.load(Ordering::SeqCst),
            1,
            "element.click() must reach the same handlers a user's click does"
        );
    }

    #[test]
    fn has_and_remove_attribute_ns_match_by_local_name() {
        let mut doc = Document::new("t");
        let d = element(&mut doc, "div", "d");
        doc.append_child(DOCUMENT, d);
        doc.set_attribute_ns(d, "http://www.w3.org/1999/xlink", "xlink:href", "/there");
        doc.set_attribute(d, "href", "/here");

        assert!(doc.has_attribute_ns(d, "http://www.w3.org/1999/xlink", "href"));
        assert!(!doc.has_attribute_ns(d, "urn:other", "href"));

        doc.remove_attribute_ns(d, "http://www.w3.org/1999/xlink", "href");
        assert!(!doc.has_attribute_ns(d, "http://www.w3.org/1999/xlink", "href"));
        assert_eq!(
            doc.get_attribute(d, "href").as_deref(),
            Some("/here"),
            "the un-namespaced attribute shares a local name and must survive"
        );
    }
}
