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
    /// The document's body and everything in it.
    ///
    /// The body IS the form, so this is the whole rendered tree. `<head>` is
    /// not synthesised: a compiled program has a title and nothing else that
    /// would go there, and inventing markup that no frontend produced would
    /// make the output stop being evidence.
    pub fn to_html(&mut self) -> String {
        let mut out = String::new();
        out.push_str("<body>");
        let children = self.child_nodes_in_order(DOCUMENT);
        if children.is_empty() {
            out.push_str("</body>");
            return out;
        }
        for child in children {
            out.push('\n');
            self.write_node(&mut out, child, 1);
        }
        out.push_str("\n</body>");
        out
    }

    /// `element.outerHTML` — the element itself and its subtree.
    pub fn outer_html(&mut self, node: NodeId) -> String {
        if node == DOCUMENT {
            return self.to_html();
        }
        let mut out = String::new();
        self.write_node(&mut out, node, 0);
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
        for (i, child) in children.into_iter().enumerate() {
            if i > 0 {
                out.push('\n');
            }
            self.write_node(&mut out, child, 0);
        }
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
        let indent = "  ".repeat(depth);
        let Some(dom_node) = self.node(node) else {
            return;
        };
        let tag = dom_node.tag.clone();

        let mut attributes: Vec<(String, String)> = dom_node
            .attributes
            .iter()
            .map(|(k, v)| (k.clone(), v.clone()))
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
        for child in children {
            out.push('\n');
            self.write_node(out, child, depth + 1);
        }
        let _ = write!(out, "\n{indent}</{tag}>");
    }
}

#[cfg(test)]
mod tests {
    use crate::dom::{DOCUMENT, Document};

    #[test]
    fn an_empty_document_is_an_empty_body() {
        let mut doc = Document::new("t");
        assert_eq!(doc.to_html(), "<body></body>");
    }

    #[test]
    fn a_button_serialises_with_its_caption_as_content() {
        let mut doc = Document::new("t");
        let b = doc.create_element("button", "");
        doc.append_child(DOCUMENT, b);
        doc.set_text_content(b, "OK");
        assert_eq!(doc.to_html(), "<body>\n  <button>OK</button>\n</body>");
    }

    #[test]
    fn a_selects_disambiguator_is_size_not_type() {
        // `<input type=checkbox>` and `<select size=6>` are the same fact
        // spelled two ways. Emitting `type` for both produced
        // `<select type="6">`, which no browser accepts — well-formed-looking
        // markup that is simply wrong, the failure mode golden files exist to
        // catch.
        let mut doc = Document::new("t");
        let list = doc.create_element("select", "6");
        doc.append_child(DOCUMENT, list);
        assert_eq!(
            doc.to_html(),
            "<body>\n  <select size=\"6\"></select>\n</body>"
        );
    }

    #[test]
    fn a_void_element_has_no_end_tag() {
        let mut doc = Document::new("t");
        let input = doc.create_element("input", "text");
        doc.append_child(DOCUMENT, input);
        assert_eq!(doc.to_html(), "<body>\n  <input type=\"text\">\n</body>");
    }

    #[test]
    fn nesting_is_indented_so_two_dumps_diff_cleanly() {
        let mut doc = Document::new("t");
        let panel = doc.create_element("div", "");
        let b = doc.create_element("button", "");
        doc.append_child(DOCUMENT, panel);
        doc.append_child(panel, b);
        doc.set_text_content(b, "Go");
        assert_eq!(
            doc.to_html(),
            "<body>\n  <div>\n    <button>Go</button>\n  </div>\n</body>"
        );
    }

    #[test]
    fn style_is_serialised_from_the_declaration_store() {
        // The reason this had to wait for the store: `color` was accepted by
        // the write side and unreadable afterwards, so it could not appear.
        let mut doc = Document::new("t");
        let b = doc.create_element("button", "");
        doc.append_child(DOCUMENT, b);
        doc.set_style_property(b, "color", "red");
        doc.set_style_property(b, "background-color", "#fff");
        assert_eq!(
            doc.to_html(),
            "<body>\n  <button style=\"background-color: #fff; color: red\"></button>\n</body>"
        );
    }

    #[test]
    fn id_leads_the_attributes_and_the_rest_are_ordered() {
        // Stable order or a golden file diffs against itself.
        let mut doc = Document::new("t");
        let b = doc.create_element("input", "text");
        doc.append_child(DOCUMENT, b);
        doc.set_attribute(b, "placeholder", "name");
        doc.set_attribute(b, "id", "field");
        assert_eq!(
            doc.to_html(),
            "<body>\n  <input id=\"field\" type=\"text\" placeholder=\"name\">\n</body>"
        );
    }

    #[test]
    fn a_boolean_attribute_serialises_as_a_bare_name() {
        let mut doc = Document::new("t");
        let b = doc.create_element("button", "");
        doc.append_child(DOCUMENT, b);
        doc.set_attribute(b, "disabled", "");
        assert_eq!(
            doc.to_html(),
            "<body>\n  <button disabled></button>\n</body>"
        );
    }

    #[test]
    fn text_is_escaped_and_so_are_attribute_values() {
        let mut doc = Document::new("t");
        let b = doc.create_element("button", "");
        doc.append_child(DOCUMENT, b);
        doc.set_text_content(b, "a < b & c");
        doc.set_attribute(b, "title", "say \"hi\"");
        assert_eq!(
            doc.to_html(),
            "<body>\n  <button title=\"say &quot;hi&quot;\">a &lt; b &amp; c</button>\n</body>"
        );
    }

    #[test]
    fn a_detached_element_is_not_in_the_document() {
        // Creating is not inserting. The whole flutter two-tree problem is a
        // created element nothing appends, and this is where that shows.
        let mut doc = Document::new("t");
        let orphan = doc.create_element("button", "");
        doc.set_text_content(orphan, "invisible");
        assert_eq!(doc.to_html(), "<body></body>");
        // …but it serialises on its own, which is what makes the bug legible.
        assert_eq!(doc.outer_html(orphan), "<button>invisible</button>");
    }

    #[test]
    fn inner_html_is_the_subtree_without_the_element() {
        let mut doc = Document::new("t");
        let panel = doc.create_element("div", "");
        let b = doc.create_element("button", "");
        doc.append_child(DOCUMENT, panel);
        doc.append_child(panel, b);
        doc.set_text_content(b, "Go");
        assert_eq!(doc.inner_html(panel), "<button>Go</button>");
        assert_eq!(doc.inner_html(b), "Go");
    }
}
