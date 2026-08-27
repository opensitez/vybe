//! `MonthCalendar` — a working calendar: it opens on this month, its arrows
//! move it, and clicking a day selects that day.
//!
//! **A calendar has no static spelling.** Every other composite control's
//! chrome is the same in every program on every day, so it travels as markup
//! (`CtorSpec::inner_html`): a `SplitContainer` has two panes, a
//! `BindingNavigator` has five items, and a string says so once. A calendar
//! does not — the month it shows is a VALUE, and it changes while the program
//! runs. So it is BUILT, through `CtorSpec::after_create`.
//!
//! ## Where the state lives: on the element, as content attributes
//!
//! `data-cal-year` / `data-cal-month` are the month being SHOWN, and
//! `selectionstart` / `selectionend` are the selection — which is not a name
//! invented here: an unmapped WinForms property lowers to `setAttribute` under
//! its own lowercased spelling, so `cal.SelectionStart` reads and writes that
//! exact attribute. **A day the user clicks and a date the program assigns are
//! therefore the same fact**, with no second copy to fall out of step.
//!
//! Nothing here adds a host function, and nothing holds interpreter-side state:
//! the calendar is `createElement` / `setAttribute` / `appendChild` /
//! `addEventListener` over the `ecma:date` leaves the rest of the .NET date
//! surface already uses — the calls a script would make, which is what keeps
//! the control swappable for a real engine.
//!
//! ## Why it re-BUILDS instead of updating in place
//!
//! `web:dom`'s `querySelector` is DOCUMENT-scoped, so a handler cannot ask
//! "the title span of THIS calendar" — two calendars on a form would both
//! answer, and the chrome deliberately carries no `id` (one must be unique per
//! document, DOM §4.9). Rebuilding needs no lookup at all: every node is in a
//! local the moment it is made. It also means construction and a month change
//! run the SAME code, so the two can never drift.
//!
//! Every host call pushes exactly one value, so each is followed by a `DROP`
//! except the one left for a caller.

use vybe_compiler::primitives::datetime::MS_PER_DAY;
use vybe_compiler::primitives::gui::{CSSOM_MODULE, DOCUMENT_MODULE, HOST_FN_ACTIVE_DOCUMENT};
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

use super::datetime_adapter;

/// The node operations live in `web:dom`; `activeDocument` is `web:html`'s.
const DOM_MODULE: &str = "web:dom";

/// The class that makes a calendar findable from inside one of its own parts.
/// `closest` is how a handler gets from the button it was given back to the
/// control — the same two calls a script would make, and the reason this is a
/// CLASS rather than an id.
const CAL_CLASS: &str = "vybe-cal";

const RENDER_FN: &str = "__dotnet_cal_render";
const PREV_FN: &str = "__dotnet_cal_prev";
const NEXT_FN: &str = "__dotnet_cal_next";
const PICK_FN: &str = "__dotnet_cal_pick";

/// The month on show.
const YEAR_ATTR: &str = "data-cal-year";
const MONTH_ATTR: &str = "data-cal-month";
/// One day cell's own date, and the month it belongs to — so clicking a
/// trailing day can move the calendar to that month without parsing anything
/// back out of the date.
const DATE_ATTR: &str = "data-cal-date";
const CELL_YEAR_ATTR: &str = "data-cal-cell-year";
const CELL_MONTH_ATTR: &str = "data-cal-cell-month";
/// ⛔ These two spellings are NOT decoration: an unmapped WinForms property
/// lowers to `setAttribute(<lowercased name>)`, so these ARE
/// `MonthCalendar.SelectionStart` / `.SelectionEnd`.
const SELECTION_START_ATTR: &str = "selectionstart";
const SELECTION_END_ATTR: &str = "selectionend";

/// Rows a month grid always has. Six, because a 31-day month beginning on the
/// last day of a week spans six of them, and a calendar that changed height
/// month to month would move every control below it.
const WEEKS: usize = 6;
const DAYS_PER_WEEK: usize = 7;

// ── Emit helpers, all on an arbitrary chunk ─────────────────────────────────

fn call_dom(chunks: &mut [Chunk], current: usize, name: &str, argc: u8, line: u32) {
    let idx = chunks[current].add_import(DOM_MODULE, name);
    chunks[current].emit_call(idx, argc, line);
}

fn call_import(chunks: &mut [Chunk], current: usize, module: &str, name: &str, argc: u8, line: u32) {
    let idx = chunks[current].add_import(module, name);
    chunks[current].emit_call(idx, argc, line);
}

/// Push the active document — the first argument of every DOM call.
fn document(chunks: &mut [Chunk], current: usize, line: u32) {
    let idx = chunks[current].add_import(DOCUMENT_MODULE, HOST_FN_ACTIVE_DOCUMENT);
    chunks[current].emit_call(idx, 0, line);
}

/// `createElement(tag)` into `slot`. Leaves the stack as it found it.
fn create_into(chunks: &mut [Chunk], current: usize, slot: u16, tag: &str, line: u32) {
    document(chunks, current, line);
    chunks[current].emit_string_const(tag, line);
    // The second argument is the input TYPE, which only an `<input>` has.
    chunks[current].emit_string_const("", line);
    call_dom(chunks, current, "createElement", 3, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, slot, line);
}

/// `node.setAttribute(name, <literal>)`, result dropped.
fn set_attribute(
    chunks: &mut [Chunk],
    current: usize,
    slot: u16,
    name: &str,
    value: &str,
    line: u32,
) {
    document(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    chunks[current].emit_string_const(name, line);
    chunks[current].emit_string_const(value, line);
    call_dom(chunks, current, "setAttribute", 4, line);
    chunks[current].emit_op(Op::DROP, line);
}

/// `node.setAttribute(name, <whatever is in `value_slot`>)`.
fn set_attribute_slot(
    chunks: &mut [Chunk],
    current: usize,
    slot: u16,
    name: &str,
    value_slot: u16,
    line: u32,
) {
    document(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    chunks[current].emit_string_const(name, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    call_dom(chunks, current, "setAttribute", 4, line);
    chunks[current].emit_op(Op::DROP, line);
}

/// Push `node.getAttribute(name)` — a string, or null when absent.
fn push_attribute(chunks: &mut [Chunk], current: usize, slot: u16, name: &str, line: u32) {
    document(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    chunks[current].emit_string_const(name, line);
    call_dom(chunks, current, "getAttribute", 3, line);
}

/// Push `node.getAttribute(name)` coerced to a NUMBER. An attribute is text;
/// the arithmetic below is not.
fn push_attribute_number(chunks: &mut [Chunk], current: usize, slot: u16, name: &str, line: u32) {
    push_attribute(chunks, current, slot, name, line);
    call_import(chunks, current, "ecma:value", "toNumber", 1, line);
}

/// One declaration of the CONTROL's own appearance.
///
/// It goes through `setStyleProperty`, so it cascades and serialises into the
/// `style` attribute like any author declaration — a later write from the
/// program simply overrides it, which is the cascade behaving normally.
fn set_style(chunks: &mut [Chunk], current: usize, slot: u16, prop: &str, value: &str, line: u32) {
    document(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    chunks[current].emit_string_const(prop, line);
    chunks[current].emit_string_const(value, line);
    call_import(chunks, current, CSSOM_MODULE, "setStyleProperty", 4, line);
    chunks[current].emit_op(Op::DROP, line);
}

/// `node.textContent = <literal>`.
fn set_text_const(chunks: &mut [Chunk], current: usize, slot: u16, text: &str, line: u32) {
    document(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    chunks[current].emit_string_const(text, line);
    call_dom(chunks, current, "setTextContent", 3, line);
    chunks[current].emit_op(Op::DROP, line);
}

/// `parent.appendChild(child)`, result dropped.
fn append(chunks: &mut [Chunk], current: usize, parent: u16, child: u16, line: u32) {
    document(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, parent, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, child, line);
    call_dom(chunks, current, "appendChild", 3, line);
    chunks[current].emit_op(Op::DROP, line);
}

/// `node.addEventListener("click", <the chunk named `handler`>)`.
///
/// The listener is an ordinary guest function value, so what a handler receives
/// is the ordinary `Event` — nothing here is a private channel.
fn on_click(
    chunks: &mut Vec<Chunk>,
    current: usize,
    slot: u16,
    handler: &str,
    line: u32,
) {
    let target = handler_chunk(chunks, handler, line);
    document(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    chunks[current].emit_string_const("click", line);
    chunks[current].emit_op_u16(Op::REF_FUNC, target as u16, line);
    chunks[current].emit(0, line); // upvalue count — the handler captures nothing
    call_dom(chunks, current, "addEventListener", 4, line);
    chunks[current].emit_op(Op::DROP, line);
}

/// Push the UTC millisecond timestamp of `start_ms + offset` whole days.
///
/// Days are added in milliseconds rather than through `DateTime.AddDays`
/// because a cell needs one NUMBER, and routing it through the date object
/// would wrap and re-read a whole `DateTime` — seven host calls — for each of
/// the forty-nine it builds.
fn push_day_ms(chunks: &mut [Chunk], current: usize, start_ms: u16, offset: usize, line: u32) {
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_GET, start_ms, line);
    if offset > 0 {
        chunk.emit_f64_const(offset as f64 * MS_PER_DAY, line);
        chunk.emit_op(Op::F64_ADD, line);
    }
}

/// Push `String(<number on the stack>)` — an attribute is text.
fn to_text(chunks: &mut [Chunk], current: usize, line: u32) {
    call_import(chunks, current, "ecma:string", "String", 1, line);
}

/// Push the `.NET`-formatted rendering of the timestamp on the stack.
fn push_formatted(chunks: &mut Vec<Chunk>, current: usize, format: &str, line: u32) {
    datetime_adapter::emit_datetime_from_millis(chunks, current, line);
    chunks[current].emit_string_const(format, line);
    datetime_adapter::emit_datetime_to_string(chunks, current, 2, line);
}

// ── The chunks ──────────────────────────────────────────────────────────────

/// Reserve a named chunk BEFORE building it, so two chunks that call each other
/// both resolve. The render function registers the handlers and every handler
/// re-renders; without reserving first that pair would recurse forever at emit
/// time instead of at run time.
fn reserve(chunks: &mut Vec<Chunk>, name: &str, arity: u8) -> (usize, bool) {
    if let Some(idx) = chunks.iter().position(|c| c.name == name) {
        return (idx, false);
    }
    let mut chunk = Chunk::new(name);
    chunk.arity = arity;
    chunk.local_count = arity as u16;
    chunks.push(chunk);
    (chunks.len() - 1, true)
}

/// Call `__dotnet_cal_render(cal)`, result dropped.
fn call_render(chunks: &mut Vec<Chunk>, current: usize, cal: u16, line: u32) {
    let target = render_chunk(chunks, line);
    chunks[current].emit_op_u16(Op::REF_FUNC, target as u16, line);
    chunks[current].emit(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cal, line);
    chunks[current].emit_op_u8_u8(Op::CALL_REF, 1, 1, line);
    chunks[current].emit_op(Op::DROP, line);
}

/// `__dotnet_cal_render(cal)` — throw away what the calendar was showing and
/// build the month named by its own attributes.
fn render_chunk(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    let (me, fresh) = reserve(chunks, RENDER_FN, 1);
    if !fresh {
        return me;
    }
    let cal: u16 = 0; // the argument
    let base = chunks[me].alloc_scratch(9);
    let (first_ms, start_ms, lead, node) = (base, base + 1, base + 2, base + 3);
    let (row, table, body, text, selected) =
        (base + 4, base + 5, base + 6, base + 7, base + 8);

    // Everything the calendar was showing, gone.
    //
    // `textContent = ""` is the DOM's own "replace all children with nothing"
    // (DOM §4.9.3: the empty string replaces all with null, adding no text
    // node), and it is one call rather than a walk.
    set_text_const(chunks, me, cal, "", line);

    // ── Which month, and which day the grid starts on ──────────────────
    //
    // `ecma:date.UTC` takes a ZERO-based month, which is the one place .NET's
    // 1-based `Month` has to be translated.
    push_attribute_number(chunks, me, cal, YEAR_ATTR, line);
    push_attribute_number(chunks, me, cal, MONTH_ATTR, line);
    chunks[me].emit_f64_const(1.0, line);
    chunks[me].emit_op(Op::F64_SUB, line);
    // day, hour, minute, second — the first of the month, at midnight.
    chunks[me].emit_f64_const(1.0, line);
    for _ in 0..3 {
        chunks[me].emit_f64_const(0.0, line);
    }
    call_import(chunks, me, "ecma:date", "UTC", 6, line);
    chunks[me].emit_op_u16(Op::LOCAL_SET, first_ms, line);

    // How many days back the grid reaches to find its Monday.
    //
    // `getUTCDay` is 0=Sunday…6=Saturday, and a WinForms calendar's default
    // week begins on Monday, so the offset is `(dow + 6) mod 7`. WASM has no
    // f64 remainder, and `x - 7*floor(x/7)` IS that remainder for the
    // non-negative values a weekday can take — with no branch.
    chunks[me].emit_op_u16(Op::LOCAL_GET, first_ms, line);
    call_import(chunks, me, "ecma:date", "getUTCDay", 1, line);
    chunks[me].emit_f64_const(6.0, line);
    chunks[me].emit_op(Op::F64_ADD, line);
    chunks[me].emit_op_u16(Op::LOCAL_SET, lead, line);
    chunks[me].emit_op_u16(Op::LOCAL_GET, lead, line);
    chunks[me].emit_op_u16(Op::LOCAL_GET, lead, line);
    chunks[me].emit_f64_const(7.0, line);
    chunks[me].emit_op(Op::F64_DIV, line);
    chunks[me].emit_op(Op::F64_FLOOR, line);
    chunks[me].emit_f64_const(7.0, line);
    chunks[me].emit_op(Op::F64_MUL, line);
    chunks[me].emit_op(Op::F64_SUB, line);
    chunks[me].emit_op_u16(Op::LOCAL_SET, lead, line);

    chunks[me].emit_op_u16(Op::LOCAL_GET, first_ms, line);
    chunks[me].emit_op_u16(Op::LOCAL_GET, lead, line);
    chunks[me].emit_f64_const(MS_PER_DAY, line);
    chunks[me].emit_op(Op::F64_MUL, line);
    chunks[me].emit_op(Op::F64_SUB, line);
    chunks[me].emit_op_u16(Op::LOCAL_SET, start_ms, line);

    // The selected date, read once — every cell is compared against it.
    push_attribute(chunks, me, cal, SELECTION_START_ATTR, line);
    chunks[me].emit_op_u16(Op::LOCAL_SET, selected, line);

    // ── The caption ────────────────────────────────────────────────────
    create_into(chunks, me, row, "div", line);
    set_attribute(chunks, me, row, "class", "vybe-cal-header", line);
    for (prop, value) in [
        ("display", "flex"),
        ("justify-content", "space-between"),
        ("align-items", "center"),
    ] {
        set_style(chunks, me, row, prop, value, line);
    }
    arrow(chunks, me, node, row, "vybe-cal-prev", "\u{25c0}", PREV_FN, line);

    create_into(chunks, me, node, "span", line);
    set_attribute(chunks, me, node, "class", "vybe-cal-title", line);
    // The caption's only elastic part: the arrows keep their glyph size and
    // the month name takes what is left.
    for (prop, value) in [("flex", "1 1 auto"), ("text-align", "center")] {
        set_style(chunks, me, node, prop, value, line);
    }
    document(chunks, me, line);
    chunks[me].emit_op_u16(Op::LOCAL_GET, node, line);
    chunks[me].emit_op_u16(Op::LOCAL_GET, first_ms, line);
    push_formatted(chunks, me, "MMMM yyyy", line);
    call_dom(chunks, me, "setTextContent", 3, line);
    chunks[me].emit_op(Op::DROP, line);
    append(chunks, me, row, node, line);

    arrow(chunks, me, node, row, "vybe-cal-next", "\u{25b6}", NEXT_FN, line);
    append(chunks, me, cal, row, line);

    // ── The grid ───────────────────────────────────────────────────────
    create_into(chunks, me, table, "table", line);
    set_attribute(chunks, me, table, "class", "vybe-cal-grid", line);
    set_style(chunks, me, table, "width", "100%", line);
    set_style(chunks, me, table, "border-collapse", "collapse", line);

    // The weekday header, read off the first week of the grid itself. Written
    // this way rather than from a list of initials because the names belong to
    // the LOCALE — the format surface already owns them, and a hardcoded
    // "M T W T F S S" would be an English calendar wearing another language's
    // month name in its own caption.
    create_into(chunks, me, body, "thead", line);
    create_into(chunks, me, row, "tr", line);
    for day in 0..DAYS_PER_WEEK {
        create_into(chunks, me, node, "th", line);
        document(chunks, me, line);
        chunks[me].emit_op_u16(Op::LOCAL_GET, node, line);
        push_day_ms(chunks, me, start_ms, day, line);
        push_formatted(chunks, me, "ddd", line);
        call_dom(chunks, me, "setTextContent", 3, line);
        chunks[me].emit_op(Op::DROP, line);
        append(chunks, me, row, node, line);
    }
    append(chunks, me, body, row, line);
    append(chunks, me, table, body, line);

    create_into(chunks, me, body, "tbody", line);
    set_attribute(chunks, me, body, "class", "vybe-cal-days", line);
    for week in 0..WEEKS {
        create_into(chunks, me, row, "tr", line);
        for day in 0..DAYS_PER_WEEK {
            let offset = week * DAYS_PER_WEEK + day;
            create_into(chunks, me, node, "td", line);
            set_attribute(chunks, me, node, "class", "vybe-cal-day", line);
            set_style(chunks, me, node, "text-align", "center", line);
            set_style(chunks, me, node, "cursor", "default", line);

            // The cell's own day number.
            document(chunks, me, line);
            chunks[me].emit_op_u16(Op::LOCAL_GET, node, line);
            push_day_ms(chunks, me, start_ms, offset, line);
            call_import(chunks, me, "ecma:date", "getUTCDate", 1, line);
            to_text(chunks, me, line);
            call_dom(chunks, me, "setTextContent", 3, line);
            chunks[me].emit_op(Op::DROP, line);

            // What the cell IS, so a click needs to parse nothing back out of
            // it: the date it selects, and the month it belongs to.
            push_day_ms(chunks, me, start_ms, offset, line);
            push_formatted(chunks, me, "yyyy-MM-dd", line);
            chunks[me].emit_op_u16(Op::LOCAL_SET, text, line);
            set_attribute_slot(chunks, me, node, DATE_ATTR, text, line);
            for (attr, getter) in [
                (CELL_YEAR_ATTR, "getUTCFullYear"),
                (CELL_MONTH_ATTR, "getUTCMonth"),
            ] {
                push_day_ms(chunks, me, start_ms, offset, line);
                call_import(chunks, me, "ecma:date", getter, 1, line);
                if getter == "getUTCMonth" {
                    // ECMAScript months are 0-based; .NET's are not, and the
                    // attribute is read back as a .NET month.
                    chunks[me].emit_f64_const(1.0, line);
                    chunks[me].emit_op(Op::F64_ADD, line);
                }
                to_text(chunks, me, line);
                chunks[me].emit_op_u16(Op::LOCAL_SET, text, line);
                set_attribute_slot(chunks, me, node, attr, text, line);
            }

            // The selected day, drawn as selected. Compared as TEXT against
            // the control's own `SelectionStart`, which is where a program's
            // assignment lands too — so a date the program set highlights
            // exactly like one the user clicked.
            push_day_ms(chunks, me, start_ms, offset, line);
            push_formatted(chunks, me, "yyyy-MM-dd", line);
            chunks[me].emit_op_u16(Op::LOCAL_GET, selected, line);
            vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[me], line);
            vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[me], line);
            chunks[me].emit_if(line);
            set_attribute(chunks, me, node, "class", "vybe-cal-day vybe-cal-selected", line);
            for (prop, value) in [("background-color", "#0a5fd8"), ("color", "#ffffff")] {
                set_style(chunks, me, node, prop, value, line);
            }
            chunks[me].emit_end(line);

            on_click(chunks, me, node, PICK_FN, line);
            append(chunks, me, row, node, line);
        }
        // The row joins the body LAST, so the table measures a row that is
        // already full rather than re-measuring once per cell.
        append(chunks, me, body, row, line);
    }
    append(chunks, me, table, body, line);
    append(chunks, me, cal, table, line);

    chunks[me].emit_op_u16(Op::LOCAL_GET, cal, line);
    chunks[me].emit_op(Op::RETURN, line);
    me
}

/// One of the caption's two month arrows, built into `slot`, wired to
/// `handler`, and appended to `parent`.
fn arrow(
    chunks: &mut Vec<Chunk>,
    current: usize,
    slot: u16,
    parent: u16,
    class: &str,
    glyph: &str,
    handler: &str,
    line: u32,
) {
    create_into(chunks, current, slot, "button", line);
    // Without an explicit type a button inside a form SUBMITS it (HTML
    // §4.10.6.1) — a calendar arrow that navigates the page is not an arrow.
    set_attribute(chunks, current, slot, "type", "button", line);
    set_attribute(chunks, current, slot, "class", class, line);
    // ⛔ The SIZE is load-bearing, not decoration. A button with no width takes
    // the toolkit's default CONTROL width, and two of those fill the caption
    // and shrink the title between them to nothing — a calendar with working
    // arrows and no month on it, which is exactly how this first rendered.
    //
    // The appearance is the CONTROL's, not HTML's: a MonthCalendar's arrows sit
    // on the title bar with no border and no fill, so both are named here
    // rather than left to the UA sheet, which draws a raised button.
    for (prop, value) in [
        ("flex", "0 0 auto"),
        ("width", "24px"),
        ("height", "20px"),
        ("padding", "0"),
        ("border", "none"),
        ("background-color", "transparent"),
        ("cursor", "default"),
    ] {
        set_style(chunks, current, slot, prop, value, line);
    }
    set_text_const(chunks, current, slot, glyph, line);
    on_click(chunks, current, slot, handler, line);
    append(chunks, current, parent, slot, line);
}

/// The chunk one of the three handlers is.
fn handler_chunk(chunks: &mut Vec<Chunk>, name: &str, line: u32) -> usize {
    match name {
        PREV_FN => step_chunk(chunks, PREV_FN, -1.0, line),
        NEXT_FN => step_chunk(chunks, NEXT_FN, 1.0, line),
        _ => pick_chunk(chunks, line),
    }
}

/// Push the calendar the event's target sits inside — `closest(".vybe-cal")`,
/// the same two calls a script would make to get from a button back to its
/// control.
///
/// The `Event` carries its target as a NODE, so this needs no lookup table and
/// no captured state: a handler is handed the document's own answer to "what
/// did the user press".
fn push_calendar_of_event(chunks: &mut [Chunk], current: usize, event: u16, line: u32) {
    document(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, event, line);
    let key = chunks[current].intern_string_constant("target");
    chunks[current].emit_struct_field_op(Op::STRUCT_GET, 0, key, line);
    chunks[current].emit_string_const(&format!(".{CAL_CLASS}"), line);
    call_dom(chunks, current, "closest", 3, line);
}

/// Push the event's target node.
fn push_event_target(chunks: &mut [Chunk], current: usize, event: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, event, line);
    let key = chunks[current].intern_string_constant("target");
    chunks[current].emit_struct_field_op(Op::STRUCT_GET, 0, key, line);
}

/// `__dotnet_cal_prev` / `__dotnet_cal_next` — move the shown month by `delta`
/// and rebuild.
///
/// The month arithmetic runs in ABSOLUTE months (`year*12 + month-1`), so a
/// December→January rollover is the same addition as any other and needs no
/// branch: `year = floor(abs/12)`, `month = abs - year*12 + 1`.
fn step_chunk(chunks: &mut Vec<Chunk>, name: &str, delta: f64, line: u32) -> usize {
    let (me, fresh) = reserve(chunks, name, 1);
    if !fresh {
        return me;
    }
    let event: u16 = 0;
    let base = chunks[me].alloc_scratch(4);
    let (cal, abs, year, text) = (base, base + 1, base + 2, base + 3);

    push_calendar_of_event(chunks, me, event, line);
    chunks[me].emit_op_u16(Op::LOCAL_SET, cal, line);

    push_attribute_number(chunks, me, cal, YEAR_ATTR, line);
    chunks[me].emit_f64_const(12.0, line);
    chunks[me].emit_op(Op::F64_MUL, line);
    push_attribute_number(chunks, me, cal, MONTH_ATTR, line);
    chunks[me].emit_op(Op::F64_ADD, line);
    chunks[me].emit_f64_const(1.0 - delta, line);
    chunks[me].emit_op(Op::F64_SUB, line);
    chunks[me].emit_op_u16(Op::LOCAL_SET, abs, line);

    chunks[me].emit_op_u16(Op::LOCAL_GET, abs, line);
    chunks[me].emit_f64_const(12.0, line);
    chunks[me].emit_op(Op::F64_DIV, line);
    chunks[me].emit_op(Op::F64_FLOOR, line);
    chunks[me].emit_op_u16(Op::LOCAL_SET, year, line);

    chunks[me].emit_op_u16(Op::LOCAL_GET, year, line);
    to_text(chunks, me, line);
    chunks[me].emit_op_u16(Op::LOCAL_SET, text, line);
    set_attribute_slot(chunks, me, cal, YEAR_ATTR, text, line);

    chunks[me].emit_op_u16(Op::LOCAL_GET, abs, line);
    chunks[me].emit_op_u16(Op::LOCAL_GET, year, line);
    chunks[me].emit_f64_const(12.0, line);
    chunks[me].emit_op(Op::F64_MUL, line);
    chunks[me].emit_op(Op::F64_SUB, line);
    chunks[me].emit_f64_const(1.0, line);
    chunks[me].emit_op(Op::F64_ADD, line);
    to_text(chunks, me, line);
    chunks[me].emit_op_u16(Op::LOCAL_SET, text, line);
    set_attribute_slot(chunks, me, cal, MONTH_ATTR, text, line);

    call_render(chunks, me, cal, line);
    chunks[me].emit_op_u16(Op::LOCAL_GET, cal, line);
    chunks[me].emit_op(Op::RETURN, line);
    me
}

/// `__dotnet_cal_pick` — the clicked day becomes the selection.
///
/// Every value it needs is already ON the cell (`data-cal-date` and the month
/// the cell belongs to), so this parses nothing. Clicking a TRAILING day moves
/// the calendar to that day's month, which is what a WinForms calendar does and
/// what makes the adjacent-month cells worth drawing at all.
fn pick_chunk(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    let (me, fresh) = reserve(chunks, PICK_FN, 1);
    if !fresh {
        return me;
    }
    let event: u16 = 0;
    let base = chunks[me].alloc_scratch(3);
    let (cal, cell, text) = (base, base + 1, base + 2);

    push_event_target(chunks, me, event, line);
    chunks[me].emit_op_u16(Op::LOCAL_SET, cell, line);
    push_calendar_of_event(chunks, me, event, line);
    chunks[me].emit_op_u16(Op::LOCAL_SET, cal, line);

    // ⛔ `SelectionStart` AND `SelectionEnd`. A single-day pick is a range of
    // one, and a program reading `SelectionRange` must not find a start with no
    // end — WinForms keeps both, and so does the control that replaces it.
    push_attribute(chunks, me, cell, DATE_ATTR, line);
    chunks[me].emit_op_u16(Op::LOCAL_SET, text, line);
    for attr in [SELECTION_START_ATTR, SELECTION_END_ATTR] {
        set_attribute_slot(chunks, me, cal, attr, text, line);
    }

    // Follow the cell into its own month.
    for (from, to) in [(CELL_YEAR_ATTR, YEAR_ATTR), (CELL_MONTH_ATTR, MONTH_ATTR)] {
        push_attribute(chunks, me, cell, from, line);
        chunks[me].emit_op_u16(Op::LOCAL_SET, text, line);
        set_attribute_slot(chunks, me, cal, to, text, line);
    }

    call_render(chunks, me, cal, line);
    chunks[me].emit_op_u16(Op::LOCAL_GET, cal, line);
    chunks[me].emit_op(Op::RETURN, line);
    me
}

// ── The construction hook ───────────────────────────────────────────────────

/// `MonthCalendar` construction — stack in `[element]`, out `[element]`.
///
/// Stamps the month the program is running in and hands over to the renderer,
/// which is the same code every later month change runs.
pub fn emit_render(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(3);
    let (cal, today, text) = (base, base + 1, base + 2);

    chunks[current].emit_op_u16(Op::LOCAL_SET, cal, line);

    // `classListAdd`, not `setAttribute("class", …)`: the control may already
    // carry classes of the program's own, and this one is only how a handler
    // finds its way back here.
    document(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cal, line);
    chunks[current].emit_string_const(CAL_CLASS, line);
    call_dom(chunks, current, "classListAdd", 3, line);
    chunks[current].emit_op(Op::DROP, line);

    datetime_adapter::emit_datetime_today(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, today, line);
    for (attr, read) in [
        (YEAR_ATTR, "Year"),
        (MONTH_ATTR, "Month"),
    ] {
        chunks[current].emit_op_u16(Op::LOCAL_GET, today, line);
        match read {
            "Year" => datetime_adapter::emit_datetime_year(chunks, current, line),
            _ => datetime_adapter::emit_datetime_month(chunks, current, line),
        }
        to_text(chunks, current, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, text, line);
        set_attribute_slot(chunks, current, cal, attr, text, line);
    }

    call_render(chunks, current, cal, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cal, line);
}
