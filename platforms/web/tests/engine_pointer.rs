//! Does a CLICK move a control's value, on a document that has been laid out?
//!
//! `dom_seam.rs` asks the same question through the seam and cannot answer it:
//! layout runs only in the paint path, so a node inserted since the last frame
//! measures 0x0 and no click can land on it. That is a property of a harness
//! that never paints, not of the running app, where a frame paints every tick.
//!
//! So this drives the engine on a document `load_html` has already laid out —
//! the state a real click actually arrives in. What it isolates is the ENGINE's
//! own answer, with the seam, the event queue and the .NET side all out of the
//! picture: if a click selects here, "the list box will not select" is a bug
//! somewhere above this line, and if it does not, this is the bug.

#![cfg(feature = "engine-webcore")]

use webcore::dom::HtmlEventType;
use webcore::types::{Document, WebCore};

const VIEWPORT: f32 = 800.0;

fn find_by_id<'a>(node: &'a WebCore, id: &str) -> Option<&'a WebCore> {
    if node.get_attr("id").as_deref() == Some(id) {
        return Some(node);
    }
    node.children.iter().find_map(|c| find_by_id(c, id))
}

/// Click a point in document coordinates, both edges, as the window shell does.
fn click(doc: &mut Document, x: f32, y: f32) {
    doc.process_mouse_event(HtmlEventType::MouseDown, (x, y), 0);
    doc.process_mouse_event(HtmlEventType::MouseUp, (x, y), 0);
}

fn rect_of(doc: &Document, id: &str) -> (f32, f32, f32, f32) {
    let n = find_by_id(&doc.root, id).unwrap_or_else(|| panic!("no #{id} in the tree"));
    let r = n.layout.content_rect;
    assert!(
        r.w > 0.0 && r.h > 0.0,
        "#{id} laid out to {}x{} — the fixture is wrong, not the engine",
        r.w,
        r.h
    );
    (r.x, r.y, r.w, r.h)
}

/// The document y of a list box row's centre.
///
/// ⛔ Off the ENGINE's own row metric, not a fraction of the box. Rows are a
/// font-derived height from the content edge, so quartering the control lands
/// on a different row as soon as the font size moves — a test that passes for
/// the wrong reason today and fails for the wrong reason tomorrow.
fn list_box_row(doc: &Document, id: &str, row: usize) -> (f32, f32) {
    let n = find_by_id(&doc.root, id).expect("no such control");
    let content = n.layout.content_rect;
    let font_px = n.style.font_size_px(16.0, 16.0).max(1.0);
    let row_h = webcore::html::forms::list_box_row_height(font_px);
    (
        content.x + content.w / 2.0,
        content.y + webcore::html::forms::LIST_BOX_PADDING + row as f32 * row_h + row_h / 2.0,
    )
}

/// A LIST BOX — `<select>` with a display size above one (HTML §4.10.7) —
/// draws its options as rows, and clicking a row selects THAT row.
///
/// This is what a WinForms `ListBox` and a VCL `TListBox` become.
#[test]
fn clicking_a_list_box_row_selects_that_row() {
    let mut doc = webcore::load_html(
        r#"<select id="lb" size="4" style="position:absolute;left:20px;top:20px;width:150px;height:100px">
             <option>one</option><option>two</option><option>three</option><option>four</option>
           </select>"#,
        VIEWPORT,
    );
    // ⛔ NOT "starts on its first option". The selectedness setting algorithm
    // auto-selects only at a display size of 1, so a list box rests with
    // NOTHING selected and `selectedIndex` is −1. The assertion this replaced
    // asserted the drop-down's rule against a list box.
    let lb = find_by_id(&doc.root, "lb").unwrap();
    assert_eq!(webcore::html::forms::selected_index(lb), -1, "a fresh list box has no selection");

    let (x, y) = list_box_row(&doc, "lb", 2);
    click(&mut doc, x, y);

    let lb = find_by_id(&doc.root, "lb").unwrap();
    assert_eq!(
        webcore::html::forms::selected_index(lb),
        2,
        "clicking the third row must select the third option"
    );
    assert_eq!(doc.open_select, 0, "a list box has no popup and must not open one");
}

/// A DROPDOWN is the same element without `size`. HTML gives `<select>` no
/// `open` IDL member, so openness is not observable in the DOM — what is
/// asserted instead is that clicking one CHANGES the selection, which is the
/// interaction a user is actually after and the one reported dead.
#[test]
fn clicking_a_dropdown_changes_its_selection() {
    let mut doc = webcore::load_html(
        r#"<select id="cb" style="position:absolute;left:20px;top:20px;width:150px;height:24px">
             <option>alpha</option><option>beta</option><option>gamma</option><option>delta</option>
           </select>"#,
        VIEWPORT,
    );
    // A drop-down DOES auto-select its first option — display size 1 is exactly
    // the case the algorithm's guard admits.
    assert_eq!(
        webcore::html::forms::selected_index(find_by_id(&doc.root, "cb").unwrap()),
        0
    );

    let (x, y, w, h) = rect_of(&doc, "cb");
    let node = find_by_id(&doc.root, "cb").unwrap().node_id;
    click(&mut doc, x + w / 2.0, y + h / 2.0);
    assert_eq!(doc.open_select, node, "the first click must OPEN the dropdown");

    // The second click lands on a row of the open list. A popup's row height is
    // the user agent's own metric with no API to ask for it, so this asserts
    // only that a click well down the list picks something OTHER than the first
    // row — enough to say the popup is live, without pinning the metric.
    click(&mut doc, x + w / 2.0, y + h + 60.0);

    assert!(
        webcore::html::forms::selected_index(find_by_id(&doc.root, "cb").unwrap()) > 0,
        "clicking down an open dropdown left the selection on the first row"
    );
    assert_eq!(doc.open_select, 0, "picking a row must close the dropdown");
}

/// A SCROLLBAR is `<input type=range>` — a value in a range — and clicking
/// along its track moves the value there. This is what `HScrollBar`,
/// `VScrollBar` and `TrackBar` become.
#[test]
fn clicking_a_range_track_moves_its_value() {
    let mut doc = webcore::load_html(
        r#"<input id="sb" type="range" min="0" max="100" step="any" value="0"
             style="position:absolute;left:20px;top:20px;width:200px;height:20px">"#,
        VIEWPORT,
    );

    let (x, y, w, h) = rect_of(&doc, "sb");
    // Three quarters along. Not the end: a thumb has width, so its centre
    // never reaches the last pixel of the track.
    click(&mut doc, x + w * 0.75, y + h / 2.0);

    // The VALUE, which is where a control's current setting lives — the `value`
    // ATTRIBUTE is its default and a click must leave it alone.
    let sb = find_by_id(&doc.root, "sb").unwrap();
    let after: f64 = webcore::html::forms::parse_floating_point(&webcore::types::input_value(sb))
        .expect("a range always holds a valid floating-point number");
    assert!(
        after > 50.0,
        "clicking three quarters along the track left the value at {after}"
    );
    assert_eq!(
        sb.get_attr("value").as_deref(),
        Some("0"),
        "a click sets the value, not the author's default"
    );
}
