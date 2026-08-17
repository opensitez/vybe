//! The user-agent stylesheet — HTML's own default rendering, as **data**.
//!
//! Every visual default a browser gives an element comes from a stylesheet, not
//! from the parser: an element arrives with a tag and attributes and nothing
//! else, and `b` is bold because a rule says so. That is the model here.
//!
//! Before this file the same defaults were imperative and scattered — a
//! `SetBordered` call for `<fieldset>`, a `SetVisible(false)` for `<dialog>`, a
//! hardcoded list inside `renders_nothing`. Each was correct and none of them
//! could be read as a stylesheet, compared against the spec, or overridden by
//! an author declaration.
//!
//! Source for every rule below:
//! <https://html.spec.whatwg.org/multipage/rendering.html>
//!
//! ## What is deliberately NOT here
//!
//! A rule earns its place by reaching something. These do not, yet, and are
//! listed so the omission is a decision rather than an oversight:
//!
//! - **`ul, ol { padding-left: 40px }`** — a `<ul>` maps to the listbox widget,
//!   which has no padding to set. It needs list markers first.
//! - **`fieldset { border: 2px groove }`** — the spec's width is 2px; we
//!   declare 1px because that is the stroke the panel actually draws. Widening
//!   the paint to match is a separate, visual decision.
//! - **`dialog:not([open]) { display: none }`** — the only rule needing an
//!   attribute condition, which would mean re-running the cascade on every
//!   attribute write. `show_dialog`/`close_dialog` already set the attribute
//!   and the visibility together.
//! (`a { color; text-decoration }` was on this list until the label grew a
//! `FontSpec`. It reaches the widget now, so it is a rule rather than an
//! excuse — which is the test for everything left above.)

/// Elements that render nothing at all.
///
/// <https://html.spec.whatwg.org/multipage/rendering.html#hidden-elements>
///
/// Only the HTML ones. A `Timer` or a `BindingSource` also renders nothing, but
/// it is a VCL/WinForms component rather than an element, and it is not the
/// stylesheet that says so.
pub const HIDDEN_ELEMENTS: &[&str] = &[
    "base", "head", "link", "meta", "noscript", "param", "script", "source", "style", "template",
    "title", "track",
];

/// One UA rule: the tags it matches, and what it declares.
type Rule = (&'static [&'static str], &'static [(&'static str, &'static str)]);

/// The rules, in cascade order. Later rules win, which is why the headings sit
/// after the generic text semantics.
const RULES: &[Rule] = &[
    // ---- Layout mode ----
    // https://html.spec.whatwg.org/multipage/rendering.html#flow-content-3
    //
    // **The property that decides what an element IS**, and until now the one
    // the stylesheet never mentioned: `display` appeared ZERO times here, so an
    // element's formatting behaviour came from `control_kind` picking a Rust
    // type at creation, where `_ => "label"` makes every unrecognised tag a
    // LEAF. A leaf takes no children, so `append_child` detaches them — which
    // is why an unknown tag does not behave the way HTML says it should.
    //
    // Declaring it changes no pixel today: the `display` arm acts on `none` and
    // treats every other value as "visible", so this is the INPUT for reading a
    // formatting context off the cascade rather than freezing it from the tag.
    //
    // Deliberately absent:
    //
    // - **`dialog`**. It is hidden at construction, before these rules are
    //   applied, so declaring any non-`none` value here would send
    //   `SetVisible(true)` and un-hide every dialog on the page. The spec's own
    //   rule is `dialog:not([open]) { display: none }`, which needs an
    //   attribute condition this table cannot express.
    // - **`list-item`**. `css::Display` has no variant for it, and inventing one
    //   that lays out like `block` would be a rule that reads as implemented and
    //   is not. The table values WERE in this list for the same reason and are
    //   now declared below — `Display` grew them along with the table
    //   formatting context that reads them.
    (
        &[
            "html", "body", "div", "p", "form", "section", "article", "main", "aside", "header",
            "footer", "nav", "hgroup", "figure", "figcaption", "blockquote", "address", "center",
            "pre", "h1", "h2", "h3", "h4", "h5", "h6", "dl", "dd", "dt", "fieldset", "legend",
            "details", "summary", "hr", "li", "ul", "ol", "menu",
        ],
        &[("display", "block")],
    ),
    (
        &[
            "span", "a", "b", "strong", "i", "em", "u", "s", "strike", "small", "big", "sub",
            "sup", "abbr", "cite", "dfn", "var", "kbd", "samp", "code", "tt", "q", "mark", "time",
            "output", "ins", "del", "bdi", "bdo", "ruby", "label",
        ],
        &[("display", "inline")],
    ),
    // A replaced element is an inline box whose CONTENT is not text — which is
    // exactly `inline-block`, and the reason an image sits in a line of prose
    // while still having a width and a height of its own.
    (
        &["img", "canvas", "iframe", "embed", "object", "input", "button", "select", "textarea",
          "progress", "meter"],
        &[("display", "inline-block")],
    ),
    // ---- Tables ----
    // https://html.spec.whatwg.org/multipage/rendering.html#tables-2
    //
    // **This is what makes `<table>` a table.** The tag used to pick a Rust
    // type — `control_kind` answered `datagridview`, so an HTML table rendered
    // as a .NET DataGrid control and its rows and cells became generic panels.
    // A table is a LAYOUT; `DataGridView` is a control; they are not the same
    // thing and the tag has no business claiming the widget.
    //
    // Each value here is read by `Formatting::Table` in `flow_layout.rs`, which
    // is the half that makes these more than labels.
    (&["table"], &[("display", "table")]),
    (&["caption"], &[("display", "table-caption"), ("text-align", "center")]),
    (&["colgroup"], &[("display", "table-column-group")]),
    (&["col"], &[("display", "table-column")]),
    (&["thead"], &[("display", "table-header-group")]),
    (&["tbody"], &[("display", "table-row-group")]),
    (&["tfoot"], &[("display", "table-footer-group")]),
    (&["tr"], &[("display", "table-row")]),
    // §14.3.9 gives cells `padding: 1px` and centres a header's text. The
    // border model is the table's and inherits down to here, which is why
    // neither is declared on the cell.
    (&["td", "th"], &[("display", "table-cell"), ("padding", "1px")]),
    (&["th"], &[("font-weight", "bold"), ("text-align", "center")]),
    // ---- The page ----
    // https://html.spec.whatwg.org/multipage/rendering.html#the-page
    //
    // `html` and `body` are the two boxes the viewport is made of, and the
    // spec gives them `height: auto` — they are as tall as their content.
    // Intrinsic sizing does not exist here (see `default_size`), so `auto`
    // means "whatever `default_size` guessed", and a parsed page would render
    // inside a 200x150 box with the rest clipped.
    //
    // `100%` is the stand-in: a page that fills its viewport is what a reader
    // sees, and unlike the guess it is derived from something real. It is
    // recorded as a stand-in rather than a rule, and it goes when boxes can
    // measure their content.
    (
        &["html", "body"],
        &[("width", "100%"), ("height", "100%")],
    ),
    // ---- Phrasing content ----
    // https://html.spec.whatwg.org/multipage/rendering.html#phrasing-content-3
    (&["b", "strong"], &[("font-weight", "bold")]),
    (
        &["cite", "dfn", "em", "i", "var", "address"],
        &[("font-style", "italic")],
    ),
    (
        &["code", "kbd", "samp", "tt", "pre"],
        &[("font-family", "monospace")],
    ),
    (&["small"], &[("font-size", "0.83em")]),
    (&["big"], &[("font-size", "1.17em")]),
    // ---- Links ----
    // https://html.spec.whatwg.org/multipage/rendering.html#the-page
    //
    // The spec's selector is `:link`/`:visited`, and a bare `<a>` with no
    // `href` is neither — it is not a link and takes no link styling. With tag
    // selectors only, this over-matches that one case. Recorded rather than
    // worked around: an attribute condition would mean re-running the cascade
    // on every attribute write, which is the same reason `dialog:not([open])`
    // is still absent below.
    // `cursor: pointer` is the third thing that makes a link look like one, and
    // it is a DECLARATION rather than widget behaviour so that an `<a>` in flow
    // — which has no widget at all, only a run — still shows a hand.
    (
        &["a"],
        &[
            ("color", "#0000ee"),
            ("text-decoration", "underline"),
            ("cursor", "pointer"),
        ],
    ),
    // ---- Text decoration ----
    (&["u", "ins"], &[("text-decoration", "underline")]),
    (
        &["s", "strike", "del"],
        &[("text-decoration", "line-through")],
    ),
    // ---- Sections and headings ----
    // https://html.spec.whatwg.org/multipage/rendering.html#sections-and-headings
    // The margins are the spec's `em` values; `font-size` is what makes an `em`
    // mean different things per level, so both are declared together.
    (
        &["h1"],
        &[
            ("font-size", "2em"),
            ("font-weight", "bold"),
            ("margin", "0.67em 0"),
        ],
    ),
    (
        &["h2"],
        &[
            ("font-size", "1.5em"),
            ("font-weight", "bold"),
            ("margin", "0.83em 0"),
        ],
    ),
    (
        &["h3"],
        &[
            ("font-size", "1.17em"),
            ("font-weight", "bold"),
            ("margin", "1em 0"),
        ],
    ),
    (
        &["h4"],
        &[
            ("font-size", "1em"),
            ("font-weight", "bold"),
            ("margin", "1.33em 0"),
        ],
    ),
    (
        &["h5"],
        &[
            ("font-size", "0.83em"),
            ("font-weight", "bold"),
            ("margin", "1.67em 0"),
        ],
    ),
    (
        &["h6"],
        &[
            ("font-size", "0.67em"),
            ("font-weight", "bold"),
            ("margin", "2.33em 0"),
        ],
    ),
    // ---- Flow content ----
    // https://html.spec.whatwg.org/multipage/rendering.html#flow-content-3
    (&["p"], &[("margin", "1em 0")]),
    (
        &["blockquote", "figure"],
        &[("margin", "1em 40px")],
    ),
    // ---- Tables ----
    // https://html.spec.whatwg.org/multipage/rendering.html#tables-2
    (&["th"], &[("font-weight", "bold"), ("text-align", "center")]),
    // ---- Form elements ----
    // https://html.spec.whatwg.org/multipage/rendering.html#form-controls
    //
    // The spec says `2px groove`; we say 1px because that is the stroke the
    // panel actually draws, and a declared border-width that does not match
    // the paint would put the containing block somewhere no pixel agrees with.
    // Widening the stroke to 2px is a separate, visual decision.
    //
    // This is the first rule where border-width is non-zero, which is the
    // first time the padding box separates from the border box: an absolutely
    // positioned child of a `<fieldset>` now starts INSIDE its frame, which is
    // both what CSS says and what VCL means by a `TGroupBox`'s client area.
    (&["fieldset"], &[("border-width", "1px")]),
    // A control's declared size is its OUTER size.
    //
    // This is the one rule below that HTML's rendering section does not
    // contain, and it is here rather than as an inverted initial value in
    // `css.rs` because it is a claim about *these elements*, not about the box
    // model. Browsers themselves disagree on the same set — Blink and WebKit
    // declare `box-sizing: border-box` on most `<input>` types and on
    // `<select>`, Gecko does so on rather fewer — so there is no single UA
    // answer to copy, and picking one is a decision to state out loud.
    //
    // The decision: an element that `control_kind` maps to a real widget takes
    // the toolkit's meaning, because that is what the frontend writing the size
    // means. `TEdit.Width = 100` is a 100px control, not a 100px interior plus
    // whatever the border came to. `<div>` and `<span>` are NOT in the list —
    // they are layout boxes rather than controls, nothing gives them a `Width`,
    // and they behave the way a browser would. Neither are the elements that
    // map to a plain label (`option`, `td`, `th`, `legend`, `summary`): those
    // are their container's CONTENT, and no frontend sizes them.
    //
    // The list therefore tracks `control_kind`, and must be extended with it.
    // `ul`/`ol` are here for a reason worth naming: they map to the listbox,
    // and the pending `ul, ol { padding-left: 40px }` marker rule would
    // otherwise turn a `TListBox.Width = 200` into 240 the day it lands.
    //
    // Inert today: only `<fieldset>` has a non-zero edge, so this changes no
    // pixel until padding arrives on the others. It is declared now so that
    // when padding does arrive, the answer is already the frontend's.
    (
        &[
            // Form controls
            "input", "button", "select", "datalist", "textarea", "fieldset", "progress", "meter",
            // Lists, tables and menus — listbox, datagridview, menustrip
            "ul", "ol", "table", "menu", // Replaced content — picturebox, canvas, linklabel
            "img", "canvas", "iframe", "embed", "object", "a",
        ],
        &[("box-sizing", "border-box")],
    ),
];

/// Whether an element renders nothing per the UA stylesheet's hidden-elements
/// rule.
pub fn is_hidden_element(tag: &str) -> bool {
    HIDDEN_ELEMENTS.contains(&tag)
}

/// Every UA declaration that applies to `tag`, in cascade order.
///
/// Tag selectors only. That is the whole selector language the UA stylesheet
/// needs for the rules above, and pretending to more would mean a selector
/// engine with one caller.
pub fn declarations_for(tag: &str) -> impl Iterator<Item = (&'static str, &'static str)> + '_ {
    RULES
        .iter()
        .filter(move |(tags, _)| tags.contains(&tag))
        .flat_map(|(_, declarations)| declarations.iter().copied())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn a_tag_with_no_rule_declares_nothing() {
        // `div` and `span` used to be the examples here. They now declare a
        // layout mode and nothing else, which is the point of `display` being
        // in the sheet at all — so the case is made with a tag the sheet has
        // genuinely never had an opinion about.
        assert_eq!(declarations_for("data").count(), 0);
        assert_eq!(declarations_for("wbr").count(), 0);
    }

    #[test]
    fn every_element_declares_the_layout_mode_it_is() {
        // The property `control_kind` used to answer by picking a Rust type.
        for (tag, mode) in [
            ("div", "block"),
            ("p", "block"),
            ("h1", "block"),
            ("li", "block"),
            ("span", "inline"),
            ("strong", "inline"),
            ("a", "inline"),
            ("label", "inline"),
            ("img", "inline-block"),
            ("input", "inline-block"),
            ("button", "inline-block"),
        ] {
            assert!(
                declarations_for(tag).any(|d| d == ("display", mode)),
                "<{tag}> must declare display: {mode}"
            );
        }
    }

    #[test]
    fn a_dialog_declares_no_layout_mode_because_it_is_hidden_before_this_runs() {
        // Not an oversight — a `display` here would reach the arm that sends
        // `SetVisible(true)` and un-hide every dialog on the page, because the
        // widget is hidden at construction and the sheet is applied after.
        assert!(
            !declarations_for("dialog").any(|d| d.0 == "display"),
            "a dialog's visibility is not this table's to state"
        );
    }

    #[test]
    fn a_control_declares_border_box_and_a_div_does_not() {
        // The toolkit convention is a DECLARATION on the elements it is true
        // of, not an inverted initial value. A `<div>` gets CSS's answer.
        for tag in [
            "input", "button", "select", "textarea", "fieldset", "ul", "ol", "table", "img", "a",
        ] {
            assert!(
                declarations_for(tag).any(|d| d == ("box-sizing", "border-box")),
                "{tag} maps to a widget: its declared size is its outer size"
            );
        }
        // A layout box, and a piece of a container's content. Neither is
        // something a frontend gives a `Width`, so both keep CSS's answer.
        for tag in ["div", "span", "td", "option"] {
            assert!(
                !declarations_for(tag).any(|d| d.0 == "box-sizing"),
                "{tag} is not a control and must keep the CSS initial value"
            );
        }
    }

    #[test]
    fn bold_and_italic_come_from_the_stylesheet_not_the_parser() {
        // Asserted by CONTAINMENT rather than as the whole list. The claim is
        // that the weight comes from a rule; pinning every declaration `strong`
        // has made this test fail the moment `display` joined the sheet, which
        // is a rule being added rather than this one breaking.
        assert!(declarations_for("strong").any(|d| d == ("font-weight", "bold")));
        assert!(declarations_for("em").any(|d| d == ("font-style", "italic")));
        // Still nothing from the parser: neither carries a weight of its own
        // that the sheet would merely be confirming.
        assert!(!declarations_for("span").any(|d| d.0 == "font-weight"));
    }

    #[test]
    fn a_heading_declares_size_weight_and_margin_together() {
        // The three are one rule because they interact: the margin is in `em`,
        // and the `em` is whatever `font-size` just made it.
        let h1: Vec<_> = declarations_for("h1").collect();
        assert!(h1.contains(&("font-size", "2em")));
        assert!(h1.contains(&("font-weight", "bold")));
        assert!(h1.contains(&("margin", "0.67em 0")));
    }

    #[test]
    fn the_hidden_elements_list_is_html_only() {
        assert!(is_hidden_element("script"));
        assert!(is_hidden_element("template"));
        // A Timer renders nothing too, but it is a component, not an element —
        // the stylesheet is not what says so.
        assert!(!is_hidden_element("vybe-timer"));
        assert!(!is_hidden_element("div"));
    }
}
