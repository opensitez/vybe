//! `New MonthCalendar()` builds a month.
//!
//! The one control whose content is a VALUE, so it is the one control a
//! declaration cannot spell: `CtorSpec::inner_html` freezes a string at compile
//! time, and a calendar frozen at compile time is wrong from the next day on.
//! It is built instead, by `platforms/dotnet`'s own adapter, through the
//! `CtorSpec::after_create` hook.
//!
//! This drives the whole path — VB source, the shared construction emit, the
//! dotnet adapter, the VM, the live document — because every layer of it is
//! new and a test of any one of them would pass while the calendar stayed
//! blank. It asserts the DOM, not the pixels: whether the caption LAYS OUT is
//! the engine's business and differs between the two, but what the tree holds
//! is the same wherever it renders.

use vybe_platform_web::engine::{DomOp, DomValue, apply};
use vybe_platform_web::html::active_document;

const SRC: &str = r#"
Imports System.Windows.Forms

Public Class Form1
    Inherits Form
End Class

Module CalendarUnderTest
    Sub Main()
        Dim f As New Form1()
        Dim cal As New MonthCalendar()
        f.Controls.Add(cal)
    End Sub
End Module
"#;

/// Compile and run the VB above, leaving its document as the active one and
/// the VM ready to receive what the user does next.
fn run_calendar_vm() -> vybe_runtime::VM {
    // ⛔ The plugin pass runs BEFORE the compile, as it does in the CLI. It is
    // what populates the language registry `normalize_class` dispatches
    // through and registers the platform trees `New MonthCalendar()` resolves
    // against — compiling first fails on both.
    // ⛔ BEFORE anything can create a document — installing after leaves the
    // one already made under the other engine.
    //
    // Pinned to webcore, and NOT because the calendar needs it. The default
    // engine's selector index is not updated when a subtree is removed, so
    // `querySelectorAll` there still matches what a redraw cleared — every
    // assertion below would count each rebuild twice and blame the control.
    // See `dom_seam::a_subtree_built_while_detached_can_still_be_cleared_once_appended`,
    // which states that as a contract and shows which engine fails it.
    vybe_platform_web::engine_select::choose(vybe_platform_web::engine_select::Engine::WebCore);
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());

    let module = vybe_language_vb::parse(SRC).expect("parse");
    let profile = vybe_compiler::profile::parse_profile(vybe_language_vb::profile_source())
        .expect("profile");
    let chunks = vybe_compiler::primitives::Compiler::with_profile(profile)
        .compile(&module)
        .expect("compile");
    vm.run(chunks).expect("run");
    vm
}

/// One turn of the runner's own loop: deliver whatever the interaction queued.
///
/// The same two calls `GuiRunner::dispatch_document_events` makes — drain the
/// document's events into (listener, event) pairs and invoke each on the VM.
/// A test that clicked and never pumped would be asserting about a handler
/// that had not run.
fn pump(vm: &mut vybe_runtime::VM) {
    let pending = vybe_platform_web::html::pending_dispatches(active_document());
    for (callback, event) in pending {
        vm.invoke(&callback, &[event]).expect("event handler");
    }
}

/// Click the centre of a node, as the window shell does: a real press and
/// release at real document coordinates, through the pointer seam.
fn click(vm: &mut vybe_runtime::VM, node: u64) {
    let DomValue::Rect { x, y, width, height } =
        apply(active_document(), DomOp::BoundingClientRect(node))
    else {
        panic!("no rect for node {node}");
    };
    assert!(
        width > 0.0 && height > 0.0,
        "nothing to click: the node laid out {width}x{height}"
    );
    let (cx, cy) = ((x + width / 2.0) as f32, (y + height / 2.0) as f32);
    for kind in ["mousedown", "mouseup"] {
        apply(
            active_document(),
            DomOp::DispatchPointer {
                kind: kind.to_string(),
                client_x: cx,
                client_y: cy,
                button: 0,
            },
        );
    }
    pump(vm);
}

/// The one node matching `selector`. Re-queried on every use ON PURPOSE: a
/// month change REBUILDS the calendar, so every node handle taken before it is
/// stale — holding one is how a test ends up asserting about a subtree that is
/// no longer in the document.
fn one(selector: &str) -> u64 {
    let found = nodes(selector);
    assert_eq!(found.len(), 1, "{selector} matched {}", found.len());
    found[0]
}

fn caption() -> String {
    text_of(one(".vybe-cal-title"))
}

fn nodes(selector: &str) -> Vec<u64> {
    match apply(
        active_document(),
        DomOp::QuerySelectorAll(selector.to_string()),
    ) {
        DomValue::Nodes(n) => n,
        other => panic!("{selector}: expected a node list, got {other:?}"),
    }
}

fn text_of(node: u64) -> String {
    match apply(active_document(), DomOp::TextContent(node)) {
        DomValue::Text(s) => s,
        other => panic!("expected text, got {other:?}"),
    }
}

fn attr_of(node: u64, name: &str) -> Option<String> {
    match apply(
        active_document(),
        DomOp::GetAttribute(node, name.to_string()),
    ) {
        DomValue::Text(s) => Some(s),
        _ => None,
    }
}

/// What a freshly constructed calendar HOLDS.
///
/// ⛔ Not a `#[test]` of its own. The document is process-global, so a second
/// test function would build a second calendar into the same tree and every
/// `.vybe-cal-title` lookup would match two — which is exactly what happened.
/// One test owns the document; this is the first half of it.
fn assert_the_grid_is_the_month_it_is_run_in() {

    // ── The grid is a month grid ───────────────────────────────────────
    let cells = nodes(".vybe-cal-day");
    assert_eq!(
        cells.len(),
        42,
        "a month grid is six weeks of seven days — built {}",
        cells.len()
    );
    assert_eq!(nodes(".vybe-cal-days tr").len(), 6, "six week rows");
    assert_eq!(
        nodes(".vybe-cal-grid thead th").len(),
        7,
        "one header per weekday"
    );

    // ── The cells are CONSECUTIVE days ─────────────────────────────────
    //
    // The invariant that does not depend on which day the suite runs: every
    // cell is the one before it plus a day, and a month boundary is the only
    // place the number may restart at 1. It catches an off-by-one in the
    // Monday alignment, a repeated cell, and a grid that stopped advancing —
    // none of which a "42 cells exist" count can see.
    let days: Vec<u32> = cells
        .iter()
        .map(|n| {
            let t = text_of(*n);
            t.trim()
                .parse()
                .unwrap_or_else(|_| panic!("a day cell must hold a day number, got {t:?}"))
        })
        .collect();
    for (i, pair) in days.windows(2).enumerate() {
        let (prev, next) = (pair[0], pair[1]);
        assert!(
            next == prev + 1 || next == 1,
            "cell {} runs {prev} → {next}: the grid must advance by a day, \
             restarting only at a month boundary",
            i + 1
        );
    }
    assert!(
        days.iter().all(|d| (1..=31).contains(d)),
        "every cell is a day of some month: {days:?}"
    );

    // ── The caption names the month ────────────────────────────────────
    let title = nodes(".vybe-cal-title");
    assert_eq!(title.len(), 1, "one title");
    let caption = text_of(title[0]);
    // The failure this catches really happened: the emit was correct and the
    // caption laid out to nothing, so the calendar rendered headless.
    let (month, year) = caption
        .rsplit_once(' ')
        .unwrap_or_else(|| panic!("the caption is `<month> <year>`, got {caption:?}"));
    assert!(
        month.chars().all(|c| c.is_alphabetic()) && !month.is_empty(),
        "the month is a NAME, not a number: {caption:?}"
    );
    assert!(
        year.len() == 4 && year.chars().all(|c| c.is_ascii_digit()),
        "the year is four digits: {caption:?}"
    );

    // ── The month arrows ───────────────────────────────────────────────
    for class in [".vybe-cal-prev", ".vybe-cal-next"] {
        let arrow = nodes(class);
        assert_eq!(arrow.len(), 1, "{class} is there exactly once");
        // ⛔ Without an explicit type a button inside a form SUBMITS it (HTML
        // §4.10.6.1), so a month arrow would navigate the page.
        assert_eq!(
            attr_of(arrow[0], "type").as_deref(),
            Some("button"),
            "{class} must not be a submit button"
        );
        assert!(
            !text_of(arrow[0]).trim().is_empty(),
            "{class} carries its glyph"
        );
    }

    // ── No `id` anywhere ───────────────────────────────────────────────
    //
    // An id must be unique per document (DOM §4.9). This template is
    // instantiated once per control, so an id here would collide the moment a
    // form held two calendars and break `getElementById` for both.
    for selector in [".vybe-cal-title", ".vybe-cal-day", ".vybe-cal-grid"] {
        for node in nodes(selector) {
            assert_eq!(
                attr_of(node, "id"),
                None,
                "{selector} must carry no id — two calendars would collide"
            );
        }
    }
}

/// The arrows MOVE the calendar, and clicking a day SELECTS it.
///
/// The point of the control. A calendar that renders the right month and does
/// nothing when you press its arrows is a picture of a calendar.
///
/// One test, like the one above and for the same reason: the document is
/// process-global, so two `#[test]`s would run concurrently against one tree.
#[test]
fn a_month_calendar_shows_the_month_navigates_and_selects() {
    let mut vm = run_calendar_vm();
    assert_the_grid_is_the_month_it_is_run_in();

    let opening = caption();
    let (month, year) = opening.rsplit_once(' ').expect("`<month> <year>`");
    let year: i32 = year.parse().expect("a four-digit year");

    // ── The arrows move it ─────────────────────────────────────────────
    click(&mut vm, one(".vybe-cal-next"));
    let next = caption();
    assert_ne!(next, opening, "the next arrow did nothing");
    click(&mut vm, one(".vybe-cal-prev"));
    assert_eq!(caption(), opening, "back is where it started");

    // ⛔ Twelve steps must land on the SAME month one year on, not on a
    // thirteenth month or a wrapped-around one. The rollover is where month
    // arithmetic goes wrong, so it is what gets counted.
    for _ in 0..12 {
        click(&mut vm, one(".vybe-cal-next"));
    }
    let (m12, y12) = {
        let c = caption();
        let (m, y) = c.rsplit_once(' ').expect("`<month> <year>`");
        (m.to_string(), y.parse::<i32>().expect("year"))
    };
    assert_eq!(m12, month, "twelve months on is the same month");
    assert_eq!(y12, year + 1, "…of the next year");
    for _ in 0..12 {
        click(&mut vm, one(".vybe-cal-prev"));
    }
    assert_eq!(caption(), opening, "and twelve back is where it started");

    // ── A day click selects that day ───────────────────────────────────
    //
    // The 15th of the shown month: never a leading or trailing cell, so it
    // cannot accidentally be testing the month-follows-the-cell rule below.
    let fifteenth = nodes(".vybe-cal-day")
        .into_iter()
        .find(|n| text_of(*n).trim() == "15" )
        .expect("a 15th");
    let date = attr_of(fifteenth, "data-cal-date").expect("the cell knows its date");
    click(&mut vm, fifteenth);

    let cal = one(".vybe-cal");
    // ⛔ `selectionstart` is not a private name — an unmapped WinForms property
    // lowers to `setAttribute` under its lowercased spelling, so this attribute
    // IS what `cal.SelectionStart` reads. A day the user clicked and a date the
    // program assigned are the same fact.
    assert_eq!(
        attr_of(cal, "selectionstart").as_deref(),
        Some(date.as_str()),
        "clicking a day must set SelectionStart"
    );
    assert_eq!(
        attr_of(cal, "selectionend").as_deref(),
        Some(date.as_str()),
        "a one-day pick is a range of one — a start with no end is not a range"
    );
    assert_eq!(caption(), opening, "picking a day inside the month does not move it");

    // And it is DRAWN as selected, which is the half a state-only assertion
    // would miss entirely.
    let selected = nodes(".vybe-cal-selected");
    assert_eq!(selected.len(), 1, "exactly one day is drawn selected");
    assert_eq!(text_of(selected[0]).trim(), "15");

    // ── A trailing day follows its own month ───────────────────────────
    //
    // The adjacent-month cells are drawn because WinForms draws them; clicking
    // one moves the calendar there, which is the only thing that makes them
    // worth drawing rather than blanking.
    let cal_month = attr_of(one(".vybe-cal"), "data-cal-month").expect("a shown month");
    let trailing = nodes(".vybe-cal-day")
        .into_iter()
        .find(|n| attr_of(*n, "data-cal-cell-month").as_deref() != Some(cal_month.as_str()))
        .expect("a leading or trailing cell");
    let trailing_month = attr_of(trailing, "data-cal-cell-month").expect("the cell knows its month");
    click(&mut vm, trailing);
    assert_eq!(
        attr_of(one(".vybe-cal"), "data-cal-month").as_deref(),
        Some(trailing_month.as_str()),
        "clicking an adjacent-month day must follow it into that month"
    );
}
