//! WinForms MonthCalendar on ordinary HTML elements.

use vybe_compiler::primitives::{ops, strings};
use vybe_runtime::Chunk;
use vybe_runtime::opcode::{Op, heaptype::HT_EXTERN};

fn document(c: &mut Chunk, line: u32) {
    let f = c.add_import("web:html", "activeDocument");
    c.emit_call(f, 0, line);
}

fn object_get(c: &mut Chunk, object: u16, key: &str, line: u32) {
    c.emit_op_u16(Op::LOCAL_GET, object, line);
    c.emit_string_const(key, line);
    let f = c.add_import("ecma:object", "get");
    c.emit_call(f, 2, line);
}

fn object_set(c: &mut Chunk, object: u16, key: &str, value: u16, line: u32) {
    c.emit_op_u16(Op::LOCAL_GET, object, line);
    c.emit_string_const(key, line);
    c.emit_op_u16(Op::LOCAL_GET, value, line);
    let f = c.add_import("ecma:object", "set");
    c.emit_call(f, 3, line);
    c.emit_op(Op::DROP, line);
}

fn style(c: &mut Chunk, node: u16, property: &str, value: &str, line: u32) {
    document(c, line);
    c.emit_op_u16(Op::LOCAL_GET, node, line);
    c.emit_string_const(property, line);
    c.emit_string_const(value, line);
    let f = c.add_import("web:cssom", "setStyleProperty");
    c.emit_call(f, 4, line);
    c.emit_op(Op::DROP, line);
}

fn bool_property(c: &mut Chunk, calendar: u16, name: &str, line: u32) -> u16 {
    object_get(c, calendar, name, line);
    let value = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, value, line);
    value
}

fn child(c: &mut Chunk, parent: u16, first: bool, line: u32) -> u16 {
    let result = c.alloc_scratch(1);
    document(c, line);
    c.emit_op_u16(Op::LOCAL_GET, parent, line);
    let f = c.add_import("web:dom", if first { "firstChild" } else { "nextSibling" });
    c.emit_call(f, 2, line);
    c.emit_op_u16(Op::LOCAL_SET, result, line);
    result
}

fn attribute(c: &mut Chunk, node: u16, name: &str, line: u32) -> u16 {
    let result = c.alloc_scratch(1);
    document(c, line);
    c.emit_op_u16(Op::LOCAL_GET, node, line);
    c.emit_string_const(name, line);
    let f = c.add_import("web:dom", "getAttribute");
    c.emit_call(f, 3, line);
    c.emit_op_u16(Op::LOCAL_SET, result, line);
    result
}

fn set_attribute(c: &mut Chunk, node: u16, name: &str, value: u16, line: u32) {
    document(c, line);
    c.emit_op_u16(Op::LOCAL_GET, node, line);
    c.emit_string_const(name, line);
    c.emit_op_u16(Op::LOCAL_GET, value, line);
    let f = c.add_import("web:dom", "setAttribute");
    c.emit_call(f, 4, line);
    c.emit_op(Op::DROP, line);
}

fn set_number(c: &mut Chunk, node: u16, name: &str, value: u16, line: u32) {
    c.emit_op_u16(Op::LOCAL_GET, value, line);
    strings::emit_to_string(c, line);
    let text = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, text, line);
    set_attribute(c, node, name, text, line);
}

fn number_attribute(c: &mut Chunk, node: u16, name: &str, line: u32) -> u16 {
    let text = attribute(c, node, name, line);
    c.emit_op_u16(Op::LOCAL_GET, text, line);
    let f = c.add_import("ecma:number", "Number");
    c.emit_call(f, 1, line);
    let value = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, value, line);
    value
}

fn date_part(c: &mut Chunk, ms: u16, part: &str, line: u32) -> u16 {
    c.emit_op_u16(Op::LOCAL_GET, ms, line);
    let f = c.add_import("ecma:date", part);
    c.emit_call(f, 1, line);
    let value = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, value, line);
    value
}

fn utc(c: &mut Chunk, year: u16, month: u16, day: f64, line: u32) -> u16 {
    c.emit_op_u16(Op::LOCAL_GET, year, line);
    c.emit_op_u16(Op::LOCAL_GET, month, line);
    c.emit_f64_const(day, line);
    let f = c.add_import("ecma:date", "UTC");
    c.emit_call(f, 3, line);
    let ms = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, ms, line);
    ms
}

fn utc_day(c: &mut Chunk, year: u16, month: u16, day: u16, line: u32) -> u16 {
    c.emit_op_u16(Op::LOCAL_GET, year, line);
    c.emit_op_u16(Op::LOCAL_GET, month, line);
    c.emit_op_u16(Op::LOCAL_GET, day, line);
    let f = c.add_import("ecma:date", "UTC");
    c.emit_call(f, 3, line);
    let ms = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, ms, line);
    ms
}

fn set_text(c: &mut Chunk, node: u16, text: u16, line: u32) {
    document(c, line);
    c.emit_op_u16(Op::LOCAL_GET, node, line);
    c.emit_op_u16(Op::LOCAL_GET, text, line);
    let f = c.add_import("web:dom", "setTextContent");
    c.emit_call(f, 3, line);
    c.emit_op(Op::DROP, line);
}

fn render(c: &mut Chunk, calendar: u16, line: u32) {
    let year = number_attribute(c, calendar, "data-year", line);
    let month = number_attribute(c, calendar, "data-month", line);
    let selected = number_attribute(c, calendar, "data-selected", line);
    let selected_year = number_attribute(c, calendar, "data-selected-year", line);
    let selected_month = number_attribute(c, calendar, "data-selected-month", line);
    let first_ms = utc(c, year, month, 1.0, line);
    let weekday = date_part(c, first_ms, "getUTCDay", line);
    c.emit_op_u16(Op::LOCAL_GET, month, line);
    c.emit_f64_const(1.0, line);
    c.emit_op(Op::F64_ADD, line);
    let next_month = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, next_month, line);
    let last_ms = utc(c, year, next_month, 0.0, line);
    let days = date_part(c, last_ms, "getUTCDate", line);

    let header = child(c, calendar, true, line);
    let previous = child(c, header, true, line);
    let title = child(c, previous, false, line);
    for name in [
        "January",
        "February",
        "March",
        "April",
        "May",
        "June",
        "July",
        "August",
        "September",
        "October",
        "November",
        "December",
    ] {
        c.emit_string_const(name, line);
    }
    c.emit_array_new_fixed(0, 12, line);
    c.emit_op_u16(Op::LOCAL_GET, month, line);
    let get = c.add_import("ecma:array", "get");
    c.emit_call(get, 2, line);
    c.emit_string_const(" ", line);
    ops::emit_dyn_add(c, line);
    c.emit_op_u16(Op::LOCAL_GET, year, line);
    strings::emit_to_string(c, line);
    ops::emit_dyn_add(c, line);
    let heading = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, heading, line);
    set_text(c, title, heading, line);

    let weekdays = child(c, header, false, line);
    let grid = child(c, weekdays, false, line);
    let footer = child(c, grid, false, line);
    let show_weeks = bool_property(c, calendar, "showweeknumbers", line);
    c.emit_op_u16(Op::LOCAL_GET, show_weeks, line);
    c.emit_if(line);
    style(c, weekdays, "grid-template-columns", "repeat(8,1fr)", line);
    style(c, grid, "grid-template-columns", "repeat(8,1fr)", line);
    c.emit_else(line);
    style(c, weekdays, "grid-template-columns", "repeat(7,1fr)", line);
    style(c, grid, "grid-template-columns", "repeat(7,1fr)", line);
    c.emit_end(line);
    let week_heading = child(c, weekdays, true, line);
    c.emit_op_u16(Op::LOCAL_GET, show_weeks, line);
    c.emit_if(line);
    style(c, week_heading, "display", "block", line);
    c.emit_else(line);
    style(c, week_heading, "display", "none", line);
    c.emit_end(line);

    object_get(c, calendar, "firstdayofweek", line);
    c.emit_string_const("Monday", line);
    ops::emit_dyn_eq(c, line);
    let monday = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, monday, line);
    let names = ["Su", "Mo", "Tu", "We", "Th", "Fr", "Sa"];
    let mut label = week_heading;
    for index in 0..7 {
        label = child(c, label, false, line);
        c.emit_op_u16(Op::LOCAL_GET, monday, line);
        c.emit_if_value(line);
        c.emit_string_const(names[(index + 1) % 7], line);
        c.emit_else(line);
        c.emit_string_const(names[index], line);
        c.emit_end(line);
        let text = c.alloc_scratch(1);
        c.emit_op_u16(Op::LOCAL_SET, text, line);
        set_text(c, label, text, line);
    }
    c.emit_op_u16(Op::LOCAL_GET, monday, line);
    c.emit_if_value(line);
    c.emit_op_u16(Op::LOCAL_GET, weekday, line);
    c.emit_f64_const(0.0, line);
    c.emit_op(Op::F64_EQ, line);
    c.emit_if_value(line);
    c.emit_f64_const(6.0, line);
    c.emit_else(line);
    c.emit_op_u16(Op::LOCAL_GET, weekday, line);
    c.emit_f64_const(1.0, line);
    c.emit_op(Op::F64_SUB, line);
    c.emit_end(line);
    c.emit_else(line);
    c.emit_op_u16(Op::LOCAL_GET, weekday, line);
    c.emit_end(line);
    let offset = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, offset, line);

    let now = c.add_import("ecma:date", "now");
    c.emit_call(now, 0, line);
    let today_ms = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, today_ms, line);
    let today_year = date_part(c, today_ms, "getUTCFullYear", line);
    let today_month = date_part(c, today_ms, "getUTCMonth", line);
    let today_day = date_part(c, today_ms, "getUTCDate", line);
    let show_today = bool_property(c, calendar, "showtoday", line);
    let circle = bool_property(c, calendar, "showtodaycircle", line);
    c.emit_op_u16(Op::LOCAL_GET, show_today, line);
    c.emit_if(line);
    style(c, footer, "display", "block", line);
    style(c, grid, "height", "calc(100% - 56px)", line);
    c.emit_string_const("Today: ", line);
    c.emit_op_u16(Op::LOCAL_GET, today_ms, line);
    let format = c.add_import("ecma:date", "toLocaleDateString");
    c.emit_call(format, 1, line);
    ops::emit_dyn_add(c, line);
    let today_text = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, today_text, line);
    set_text(c, footer, today_text, line);
    c.emit_else(line);
    style(c, footer, "display", "none", line);
    style(c, grid, "height", "calc(100% - 38px)", line);
    c.emit_end(line);

    let mut cursor = child(c, grid, true, line);
    for row in 0..6 {
        if row != 0 {
            cursor = child(c, cursor, false, line);
        }
        let marker = cursor;
        c.emit_op_u16(Op::LOCAL_GET, show_weeks, line);
        c.emit_if(line);
        style(c, marker, "display", "block", line);
        c.emit_f64_const(row as f64 * 7.0 + 1.0, line);
        c.emit_op_u16(Op::LOCAL_GET, offset, line);
        c.emit_op(Op::F64_SUB, line);
        let start_day = c.alloc_scratch(1);
        c.emit_op_u16(Op::LOCAL_SET, start_day, line);
        let start_ms = utc_day(c, year, month, start_day, line);
        let week_year = date_part(c, start_ms, "getUTCFullYear", line);
        c.emit_f64_const(0.0, line);
        let january = c.alloc_scratch(1);
        c.emit_op_u16(Op::LOCAL_SET, january, line);
        let january_first = utc(c, week_year, january, 1.0, line);
        c.emit_op_u16(Op::LOCAL_GET, start_ms, line);
        c.emit_op_u16(Op::LOCAL_GET, january_first, line);
        c.emit_op(Op::F64_SUB, line);
        c.emit_f64_const(604_800_000.0, line);
        c.emit_op(Op::F64_DIV, line);
        c.emit_op(Op::I32_FROM_F64, line);
        c.emit_i32_const(1, line);
        c.emit_op(Op::I32_ADD, line);
        strings::emit_to_string(c, line);
        let week_text = c.alloc_scratch(1);
        c.emit_op_u16(Op::LOCAL_SET, week_text, line);
        set_text(c, marker, week_text, line);
        c.emit_else(line);
        style(c, marker, "display", "none", line);
        c.emit_end(line);

        for column in 0..7 {
            cursor = child(c, cursor, false, line);
            let button = cursor;
            c.emit_f64_const(row as f64 * 7.0 + column as f64 + 1.0, line);
            c.emit_op_u16(Op::LOCAL_GET, offset, line);
            c.emit_op(Op::F64_SUB, line);
            let day = c.alloc_scratch(1);
            c.emit_op_u16(Op::LOCAL_SET, day, line);
            c.emit_op_u16(Op::LOCAL_GET, day, line);
            c.emit_f64_const(1.0, line);
            c.emit_op(Op::F64_GE, line);
            c.emit_op_u16(Op::LOCAL_GET, day, line);
            c.emit_op_u16(Op::LOCAL_GET, days, line);
            c.emit_op(Op::F64_LE, line);
            c.emit_op(Op::I32_AND, line);
            c.emit_if(line);
            set_number(c, button, "data-day", day, line);
            c.emit_op_u16(Op::LOCAL_GET, day, line);
            strings::emit_to_string(c, line);
            let text = c.alloc_scratch(1);
            c.emit_op_u16(Op::LOCAL_SET, text, line);
            set_text(c, button, text, line);
            c.emit_op_u16(Op::LOCAL_GET, day, line);
            c.emit_op_u16(Op::LOCAL_GET, selected, line);
            c.emit_op(Op::F64_EQ, line);
            c.emit_op_u16(Op::LOCAL_GET, year, line);
            c.emit_op_u16(Op::LOCAL_GET, selected_year, line);
            c.emit_op(Op::F64_EQ, line);
            c.emit_op(Op::I32_AND, line);
            c.emit_op_u16(Op::LOCAL_GET, month, line);
            c.emit_op_u16(Op::LOCAL_GET, selected_month, line);
            c.emit_op(Op::F64_EQ, line);
            c.emit_op(Op::I32_AND, line);
            let selected_day = c.alloc_scratch(1);
            c.emit_op_u16(Op::LOCAL_SET, selected_day, line);
            c.emit_op_u16(Op::LOCAL_GET, circle, line);
            c.emit_op_u16(Op::LOCAL_GET, day, line);
            c.emit_op_u16(Op::LOCAL_GET, today_day, line);
            c.emit_op(Op::F64_EQ, line);
            c.emit_op(Op::I32_AND, line);
            c.emit_op_u16(Op::LOCAL_GET, year, line);
            c.emit_op_u16(Op::LOCAL_GET, today_year, line);
            c.emit_op(Op::F64_EQ, line);
            c.emit_op(Op::I32_AND, line);
            c.emit_op_u16(Op::LOCAL_GET, month, line);
            c.emit_op_u16(Op::LOCAL_GET, today_month, line);
            c.emit_op(Op::F64_EQ, line);
            c.emit_op(Op::I32_AND, line);
            let today_circle = c.alloc_scratch(1);
            c.emit_op_u16(Op::LOCAL_SET, today_circle, line);
            c.emit_op_u16(Op::LOCAL_GET, selected_day, line);
            c.emit_if_value(line);
            c.emit_op_u16(Op::LOCAL_GET, today_circle, line);
            c.emit_if_value(line);
            c.emit_string_const("background:#2274bf;color:white;border:0;outline:1px solid #cf3b35;padding:0;font-size:11px;min-width:0", line);
            c.emit_else(line);
            c.emit_string_const("background:#2274bf;color:white;border:0;padding:0;font-size:11px;min-width:0", line);
            c.emit_end(line);
            c.emit_else(line);
            c.emit_op_u16(Op::LOCAL_GET, today_circle, line);
            c.emit_if_value(line);
            c.emit_string_const("background:transparent;color:inherit;border:0;outline:1px solid #cf3b35;padding:0;font-size:11px;min-width:0", line);
            c.emit_else(line);
            c.emit_string_const("background:transparent;color:inherit;border:0;padding:0;font-size:11px;min-width:0", line);
            c.emit_end(line);
            c.emit_end(line);
            let day_style = c.alloc_scratch(1);
            c.emit_op_u16(Op::LOCAL_SET, day_style, line);
            set_attribute(c, button, "style", day_style, line);
            c.emit_else(line);
            c.emit_string_const("", line);
            let blank = c.alloc_scratch(1);
            c.emit_op_u16(Op::LOCAL_SET, blank, line);
            set_text(c, button, blank, line);
            set_attribute(c, button, "data-day", blank, line);
            c.emit_string_const("background:transparent;color:inherit;border:0;padding:0;font-size:11px;min-width:0", line);
            let blank_style = c.alloc_scratch(1);
            c.emit_op_u16(Op::LOCAL_SET, blank_style, line);
            set_attribute(c, button, "style", blank_style, line);
            c.emit_end(line);
        }
    }
}

fn event_node(c: &mut Chunk, key: &str, line: u32) -> u16 {
    c.emit_op_u16(Op::LOCAL_GET, 0, line);
    c.emit_string_const(key, line);
    let get = c.add_import("ecma:object", "get");
    c.emit_call(get, 2, line);
    let id = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, id, line);
    let new = c.add_import("ecma:object", "new");
    c.emit_call(new, 0, line);
    let node = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, node, line);
    c.emit_op_u16(Op::LOCAL_GET, node, line);
    c.emit_string_const("__node", line);
    c.emit_op_u16(Op::LOCAL_GET, id, line);
    let set = c.add_import("ecma:object", "set");
    c.emit_call(set, 3, line);
    c.emit_op(Op::DROP, line);
    node
}

fn advance(c: &mut Chunk, calendar: u16, direction: f64, line: u32) {
    let year = number_attribute(c, calendar, "data-year", line);
    let month = number_attribute(c, calendar, "data-month", line);
    c.emit_op_u16(Op::LOCAL_GET, month, line);
    object_get(c, calendar, "scrollchange", line);
    let number = c.add_import("ecma:number", "Number");
    c.emit_call(number, 1, line);
    c.emit_f64_const(direction, line);
    c.emit_op(Op::F64_MUL, line);
    c.emit_op(Op::F64_ADD, line);
    let moved_month = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, moved_month, line);
    let ms = utc(c, year, moved_month, 1.0, line);
    let new_year = date_part(c, ms, "getUTCFullYear", line);
    let new_month = date_part(c, ms, "getUTCMonth", line);
    set_number(c, calendar, "data-year", new_year, line);
    set_number(c, calendar, "data-month", new_month, line);
    render(c, calendar, line);
}

fn callback(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    let mut c = vybe_compiler::primitives::functions::create_function_chunk(
        "__dotnet_monthcalendar_click",
        1,
    );
    c.local_count = 1;
    let calendar = event_node(&mut c, "currentTarget", line);
    let target = event_node(&mut c, "target", line);
    let action = attribute(&mut c, target, "data-action", line);
    c.emit_op_u16(Op::LOCAL_GET, action, line);
    c.emit_string_const("prev", line);
    ops::emit_dyn_eq(&mut c, line);
    c.emit_if(line);
    advance(&mut c, calendar, -1.0, line);
    c.emit_else(line);
    c.emit_op_u16(Op::LOCAL_GET, action, line);
    c.emit_string_const("next", line);
    ops::emit_dyn_eq(&mut c, line);
    c.emit_if(line);
    advance(&mut c, calendar, 1.0, line);
    c.emit_else(line);
    let day_text = attribute(&mut c, target, "data-day", line);
    c.emit_op_u16(Op::LOCAL_GET, day_text, line);
    let truthy = c.add_import("ecma:boolean", "toBoolean");
    c.emit_call(truthy, 1, line);
    c.emit_if(line);
    set_attribute(&mut c, calendar, "data-selected", day_text, line);
    let year = attribute(&mut c, calendar, "data-year", line);
    let month = attribute(&mut c, calendar, "data-month", line);
    set_attribute(&mut c, calendar, "data-selected-year", year, line);
    set_attribute(&mut c, calendar, "data-selected-month", month, line);
    render(&mut c, calendar, line);
    c.emit_end(line);
    c.emit_end(line);
    c.emit_end(line);
    c.emit_ref_null(HT_EXTERN, line);
    c.emit_op(Op::RETURN, line);
    chunks.push(c);
    chunks.len() - 1
}

/// Stack: [calendar] -> [null].
pub fn emit_init(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let c = &mut chunks[current];
    let calendar = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, calendar, line);
    for (key, value) in [
        ("showtoday", true),
        ("showtodaycircle", true),
        ("showweeknumbers", false),
    ] {
        c.emit_bool_const(value, line);
        let slot = c.alloc_scratch(1);
        c.emit_op_u16(Op::LOCAL_SET, slot, line);
        object_set(c, calendar, key, slot, line);
    }
    for (key, value) in [("maxselectioncount", 7), ("scrollchange", 1)] {
        c.emit_i32_const(value, line);
        let slot = c.alloc_scratch(1);
        c.emit_op_u16(Op::LOCAL_SET, slot, line);
        object_set(c, calendar, key, slot, line);
    }
    c.emit_string_const("Default", line);
    let first_day = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, first_day, line);
    object_set(c, calendar, "firstdayofweek", first_day, line);
    let now = c.add_import("ecma:date", "now");
    c.emit_call(now, 0, line);
    let ms = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, ms, line);
    let year = date_part(c, ms, "getUTCFullYear", line);
    let month = date_part(c, ms, "getUTCMonth", line);
    let day = date_part(c, ms, "getUTCDate", line);
    set_number(c, calendar, "data-year", year, line);
    set_number(c, calendar, "data-month", month, line);
    set_number(c, calendar, "data-selected", day, line);
    set_number(c, calendar, "data-selected-year", year, line);
    set_number(c, calendar, "data-selected-month", month, line);
    render(c, calendar, line);
    let handler = callback(chunks, line);
    vybe_compiler::primitives::gui::emit_add_event_listener(
        &mut chunks[current],
        calendar,
        "click",
        handler,
        line,
    );
    chunks[current].emit_ref_null(HT_EXTERN, line);
}

/// Stack: [calendar] -> [property value].
pub fn emit_property_get(c: &mut Chunk, property: &str, line: u32) {
    let calendar = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, calendar, line);
    object_get(c, calendar, property, line);
}

/// Stack: [calendar, value] -> [null].
pub fn emit_property_set(c: &mut Chunk, property: &str, line: u32) {
    let value = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, value, line);
    let calendar = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, calendar, line);
    object_set(c, calendar, property, value, line);
    if matches!(
        property,
        "showtoday" | "showtodaycircle" | "showweeknumbers" | "firstdayofweek"
    ) {
        render(c, calendar, line);
    }
    c.emit_ref_null(HT_EXTERN, line);
}
