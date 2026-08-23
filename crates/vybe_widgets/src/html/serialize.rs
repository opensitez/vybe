//! The live document, written back out as HTML.
//!
//! `element.outerHTML` over the widget tree. Not a debug dump that resembles
//! markup — the real tag, the real attributes, the real inline style, indented
//! so two of them diff cleanly.
//!
//! ## Why this earns its place
//!
//! **It is the only automatic check that the GUI renders anything.** The GUI
//! corpus does not cover the render path: two pascal GUI slices are broken
//! extraction, and no flutter test reaches `runApp` at all. Every verification
//! so far has been a person looking at a PNG. A serialised tree is diffable,
//! reviewable in a code review, and fails loudly when a control lands in the
//! wrong parent or with the wrong tag.
//!
//! **It measures the goal directly.** The target for every frontend is to
//! *become HTML*. If the output here is not markup you would have written by
//! hand for the same UI, the mapping is wrong — a `<div>` where `<header>`
//! belongs, or a `vybe-menustrip` that should have been `<menu>`, is visible
//! here and invisible in a screenshot.
//!
//! This became possible only once declarations were stored
//! ([`crate::css::Style`]). Before that a style write was consumed into a
//! widget command and forgotten, so there was nothing to serialise: `color`
//! and `font-size` had been *set* and could not be read back.
//!
//! ## What it reads, and from where
//!
//! Each fact comes from wherever it actually lives — there is no second copy to
//! serialise from:
//!
//! - **tag / `type`** — the node, as `createElement` built it.
//! - **attributes** — the node's attribute map. Note this is HTML-correct
//!   rather than convenient: `value` is the *content attribute* (the default),
//!   not the live IDL value, which is exactly what a browser serialises.
//! - **text** — the control, via `GetText`.
//! - **style** — the declaration store, verbatim.

use std::fmt::Write as _;

use crate::dom::{DOCUMENT, Document, NodeId};

/// Elements with no end tag and no children (HTML Standard §13.1.2).
const VOID: &[&str] = &[
    "area", "base", "br", "col", "embed", "hr", "img", "input", "link", "meta", "param", "source",
    "track", "wbr",
];

fn is_void(tag: &str) -> bool {
    VOID.contains(&tag)
}

/// Attributes worth emitting first, in this order, because they identify the
/// element. Everything else follows alphabetically so output is stable.
const LEADING: &[&str] = &["id", "class", "type", "name"];

pub fn escape_text(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
}

pub fn escape_attribute(s: &str) -> String {
    escape_text(s).replace('"', "&quot;")
}

impl Document {
    /// The document, serialised — `<html>`, `<head>`, `<body>` and everything
    /// in them.
    ///
    /// The wrapper used to be SYNTHESISED here: a literal `"<body>"` was
    /// written around the document's children, because the document node was
    /// standing in for the body. Now that tree construction puts real elements
    /// in the tree ([`Document::body`]), writing one as well produced a body
    /// inside a body. Serialising the children is both simpler and more
    /// honest — the output is the tree, with nothing invented, which is the
    /// only reason this is usable as evidence.
    pub fn to_html(&mut self) -> String {
        let mut out = String::new();
        let children = self.child_nodes_in_order(DOCUMENT);
        for (index, child) in children.into_iter().enumerate() {
            if index > 0 {
                out.push('\n');
            }
            self.write_node(&mut out, child, 0);
        }
        out
    }

    /// `element.outerHTML` — the element itself and its subtree.
    ///
    /// **No added whitespace.** The HTML fragment serialization algorithm
    /// writes tags and content and nothing else, and the difference is not
    /// cosmetic: `outerHTML` is a round trip, and re-parsing indentation turns
    /// it into TEXT NODES the original tree never had. `to_html` is the
    /// pretty one because it is a document dump for a reader, not markup
    /// anyone parses back.
    pub fn outer_html(&mut self, node: NodeId) -> String {
        if node == DOCUMENT {
            return self.to_html();
        }
        let mut out = String::new();
        self.write_node_compact(&mut out, node);
        out
    }

    /// `element.innerHTML` — the subtree, without the element itself.
    ///
    /// A text-bearing control's text IS its content, so a `<button>` answers
    /// its caption here exactly as a browser would.
    pub fn inner_html(&mut self, node: NodeId) -> String {
        let children = self.child_nodes_in_order(node);
        if children.is_empty() {
            return escape_text(&self.text_content(node));
        }
        let mut out = String::new();
        for child in children {
            // No separator: see `outer_html`. A newline here is markup the
            // tree does not contain.
            self.write_node_compact(&mut out, child);
        }
        out
    }

    /// [`write_node`](Self::write_node) with indentation and line breaks
    /// suppressed — the fragment serialization the DOM getters owe.
    fn write_node_compact(&mut self, out: &mut String, node: NodeId) {
        self.write_node_indented(out, node, 0, false);
    }

    /// One node's subtree, laid out for a READER.
    ///
    /// The same thing [`to_html`](Self::to_html) does for the whole document,
    /// and deliberately NOT what [`outer_html`](Self::outer_html) does: an
    /// inspector wants the indentation, and a DOM getter must not invent it.
    /// The debugger's node dump is the caller.
    pub fn outer_html_pretty(&mut self, node: NodeId) -> String {
        if node == DOCUMENT {
            return self.to_html();
        }
        let mut out = String::new();
        self.write_node(&mut out, node, 0);
        out
    }

    /// A node's children, in document order.
    ///
    /// Derived from `parent` pointers rather than read off a child list,
    /// because the document element is not itself a node — its children record
    /// their parent, but there is no `DomNode` holding them. Creation order is
    /// document order, which is the same rule `elements()` states.
    fn child_nodes_in_order(&self, parent: NodeId) -> Vec<NodeId> {
        self.elements()
            .into_iter()
            .filter(|id| self.node(*id).and_then(|n| n.parent) == Some(parent))
            .collect()
    }

    fn write_node(&mut self, out: &mut String, node: NodeId, depth: usize) {
        self.write_node_indented(out, node, depth, true);
    }

    /// The one serialiser. `pretty` decides whether it lays the tree out for a
    /// reader or writes the fragment markup the DOM getters owe — see
    /// [`outer_html`](Self::outer_html).
    fn write_node_indented(
        &mut self,
        out: &mut String,
        node: NodeId,
        depth: usize,
        pretty: bool,
    ) {
        let indent = if pretty { "  ".repeat(depth) } else { String::new() };
        let Some(dom_node) = self.node(node) else {
            return;
        };
        // A node that is not an element has no tag, no attributes and no end
        // tag — it serialises as its data. Falling through to the element path
        // wrote `<>…</>` for both, which is not markup at all: the tree could
        // hold a text node and a comment long before it could write one back.
        match dom_node.kind {
            crate::dom::NodeKind::Text => {
                let _ = write!(out, "{indent}{}", escape_text(&dom_node.data));
                return;
            }
            // Comment data is NOT escaped: `<!--` … `-->` is a raw run, and
            // `&amp;` inside one is four characters, not one.
            crate::dom::NodeKind::Comment => {
                let _ = write!(out, "{indent}<!--{}-->", dom_node.data);
                return;
            }
            // XML-only productions, serialised as themselves. Neither escapes
            // its data: `<![CDATA[` … `]]>` and `<?` … `?>` are raw runs, and
            // an `&amp;` inside one is five characters rather than one.
            crate::dom::NodeKind::CData => {
                let _ = write!(out, "{indent}<![CDATA[{}]]>", dom_node.data);
                return;
            }
            crate::dom::NodeKind::ProcessingInstruction => {
                let data = &dom_node.data;
                let _ = if data.is_empty() {
                    write!(out, "{indent}<?{}?>", dom_node.tag)
                } else {
                    write!(out, "{indent}<?{} {}?>", dom_node.tag, data)
                };
                return;
            }
            // A fragment has no markup of its own — DOM Parsing serialises a
            // fragment AS its children, which is the same thing `innerHTML`
            // asks for. Writing a tag here would invent one that no parser
            // could read back.
            crate::dom::NodeKind::DocumentFragment => {
                // Children collected BEFORE recursing, because `write_node`
                // takes `&mut self` and the borrow of `dom_node` cannot
                // outlive it.
                for child in self.child_nodes_in_order(node) {
                    self.write_node_indented(out, child, depth, pretty);
                }
                return;
            }
            crate::dom::NodeKind::Element => {}
        }
        let tag = dom_node.tag.clone();

        let mut attributes: Vec<(String, String)> = dom_node
            .attributes
            .iter()
            .map(|a| (a.name.clone(), a.value.clone()))
            .collect();
        // The create-time disambiguator is not in the attribute map — it is a
        // field on the node, because with the tag it is what decides which
        // control this is. Omitting it would serialise every input as a text
        // field: HTML's missing-value default makes that *valid* markup and the
        // wrong control, which is the worst combination in a golden file.
        //
        // WHICH attribute it is depends on the element. `<input type=checkbox>`
        // and `<select size=6>` are the same fact spelled two ways, and
        // `<select type="6">` is not markup a browser would accept — so the
        // serialiser has to ask the tag, not assume `type`.
        let disambiguator = match dom_node.tag.as_str() {
            "select" | "datalist" => "size",
            _ => "type",
        };
        if !dom_node.input_type.is_empty()
            && !attributes.iter().any(|(name, _)| name == disambiguator)
        {
            attributes.push((disambiguator.to_string(), dom_node.input_type.clone()));
        }
        attributes.sort_by(|a, b| {
            let rank = |k: &str| {
                LEADING
                    .iter()
                    .position(|l| *l == k)
                    .unwrap_or(LEADING.len())
            };
            rank(&a.0).cmp(&rank(&b.0)).then_with(|| a.0.cmp(&b.0))
        });

        let style = self
            .style(node)
            .filter(|s| !s.is_empty())
            .map(|s| s.css_text());

        let _ = write!(out, "{indent}<{tag}");
        for (name, value) in attributes {
            if value.is_empty() {
                // A boolean content attribute: presence is truth, so the bare
                // name is the correct serialisation.
                let _ = write!(out, " {name}");
            } else {
                let _ = write!(out, " {name}=\"{}\"", escape_attribute(&value));
            }
        }
        if let Some(style) = style {
            let _ = write!(out, " style=\"{}\"", escape_attribute(&style));
        }
        out.push('>');

        if is_void(&tag) {
            return;
        }

        let children = self.child_nodes_in_order(node);
        if children.is_empty() {
            let text = self.text_content(node);
            let _ = write!(out, "{}</{tag}>", escape_text(&text));
            return;
        }
        // **An element whose only child is text stays on ONE line.**
        //
        // Indentation would put whitespace INSIDE the element, and whitespace
        // in HTML is content: `<button>\n  OK\n</button>` is a different button
        // from `<button>OK</button>`, and a dump that cannot be read back is
        // not a dump of this tree. Browsers' inspectors make the same
        // exception, for the same reason.
        //
        // This used to fall out of the `children.is_empty()` branch above,
        // because a caption was a STRING on the element rather than a node.
        // Now that `textContent` inserts a real Text node — as DOM §4.4 says it
        // must — the caption is a child and the rule has to be stated.
        let inline_text = children.len() == 1 && self.is_text_node(children[0]);
        for child in children {
            if pretty && !inline_text {
                out.push('\n');
            }
            // `pretty` off for the inline case: a text node writes its own
            // indent prefix, so suppressing only the newline would leave the
            // spaces behind — `<button>      OK</button>`.
            self.write_node_indented(out, child, depth + 1, pretty && !inline_text);
        }
        if pretty && !inline_text {
            let _ = write!(out, "\n{indent}");
        }
        let _ = write!(out, "</{tag}>");
    }
}

#[cfg(test)]
mod tests {
    use crate::dom::{DOCUMENT, Document};

    /// The document a browser serialises around any body content.
    ///
    /// `<html>`, `<head>` and `<body>` are always in the tree — tree
    /// construction inserts them whether or not the markup said so — so every
    /// expectation here carries them. Written once rather than ten times, so
    /// what each test is actually about stays legible.
    fn page(body: &str) -> String {
        if body.is_empty() {
            return "<html>\n  <head></head>\n  <body></body>\n</html>".to_string();
        }
        format!("<html>\n  <head></head>\n  <body>\n{body}\n  </body>\n</html>")
    }

    #[test]
    fn an_empty_document_is_an_empty_body() {
        let mut doc = Document::new("t");
        assert_eq!(doc.to_html(), page(""));
    }

    #[test]
    fn a_button_serialises_with_its_caption_as_content() {
        let mut doc = Document::new("t");
        let b = doc.create_element_typed("button", "");
        doc.append_child(DOCUMENT, b);
        doc.set_text_content(b, "OK");
        assert_eq!(doc.to_html(), page("    <button>OK</button>"));
    }

    #[test]
    fn a_selects_disambiguator_is_size_not_type() {
        // `<input type=checkbox>` and `<select size=6>` are the same fact
        // spelled two ways. Emitting `type` for both produced
        // `<select type="6">`, which no browser accepts — well-formed-looking
        // markup that is simply wrong, the failure mode golden files exist to
        // catch.
        let mut doc = Document::new("t");
        let list = doc.create_element_typed("select", "6");
        doc.append_child(DOCUMENT, list);
        assert_eq!(doc.to_html(), page("    <select size=\"6\"></select>"));
    }

    #[test]
    fn a_void_element_has_no_end_tag() {
        let mut doc = Document::new("t");
        let input = doc.create_element_typed("input", "text");
        doc.append_child(DOCUMENT, input);
        assert_eq!(doc.to_html(), page("    <input type=\"text\">"));
    }

    #[test]
    fn nesting_is_indented_so_two_dumps_diff_cleanly() {
        let mut doc = Document::new("t");
        let panel = doc.create_element_typed("div", "");
        let b = doc.create_element_typed("button", "");
        doc.append_child(DOCUMENT, panel);
        doc.append_child(panel, b);
        doc.set_text_content(b, "Go");
        assert_eq!(
            doc.to_html(),
            page("    <div>\n      <button>Go</button>\n    </div>")
        );
    }

    #[test]
    fn style_is_serialised_from_the_declaration_store() {
        // The reason this had to wait for the store: `color` was accepted by
        // the write side and unreadable afterwards, so it could not appear.
        let mut doc = Document::new("t");
        let b = doc.create_element_typed("button", "");
        doc.append_child(DOCUMENT, b);
        doc.set_style_property(b, "color", "red");
        doc.set_style_property(b, "background-color", "#fff");
        assert_eq!(
            doc.to_html(),
            page("    <button style=\"background-color: #fff; color: red\"></button>")
        );
    }

    #[test]
    fn id_leads_the_attributes_and_the_rest_are_ordered() {
        // Stable order or a golden file diffs against itself.
        let mut doc = Document::new("t");
        let b = doc.create_element_typed("input", "text");
        doc.append_child(DOCUMENT, b);
        doc.set_attribute(b, "placeholder", "name");
        doc.set_attribute(b, "id", "field");
        assert_eq!(
            doc.to_html(),
            page("    <input id=\"field\" type=\"text\" placeholder=\"name\">")
        );
    }

    #[test]
    fn a_boolean_attribute_serialises_as_a_bare_name() {
        let mut doc = Document::new("t");
        let b = doc.create_element_typed("button", "");
        doc.append_child(DOCUMENT, b);
        doc.set_attribute(b, "disabled", "");
        assert_eq!(
            doc.to_html(),
            page("    <button disabled></button>")
        );
    }

    #[test]
    fn text_is_escaped_and_so_are_attribute_values() {
        let mut doc = Document::new("t");
        let b = doc.create_element_typed("button", "");
        doc.append_child(DOCUMENT, b);
        doc.set_text_content(b, "a < b & c");
        doc.set_attribute(b, "title", "say \"hi\"");
        assert_eq!(
            doc.to_html(),
            page("    <button title=\"say &quot;hi&quot;\">a &lt; b &amp; c</button>")
        );
    }

    #[test]
    fn a_detached_element_is_not_in_the_document() {
        // Creating is not inserting. The whole flutter two-tree problem is a
        // created element nothing appends, and this is where that shows.
        let mut doc = Document::new("t");
        let orphan = doc.create_element_typed("button", "");
        doc.set_text_content(orphan, "invisible");
        assert_eq!(doc.to_html(), page(""));
        // …but it serialises on its own, which is what makes the bug legible.
        assert_eq!(doc.outer_html(orphan), "<button>invisible</button>");
    }

    #[test]
    fn inner_html_is_the_subtree_without_the_element() {
        let mut doc = Document::new("t");
        let panel = doc.create_element_typed("div", "");
        let b = doc.create_element_typed("button", "");
        doc.append_child(DOCUMENT, panel);
        doc.append_child(panel, b);
        doc.set_text_content(b, "Go");
        assert_eq!(doc.inner_html(panel), "<button>Go</button>");
        assert_eq!(doc.inner_html(b), "Go");
    }
}
