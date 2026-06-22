use vybe_bytecode::Chunk;

pub fn split(c: &mut Chunk, line: u32) {
    let idx = c.add_import("ecma:string", "split");
    c.emit_call(idx, 2, line);
}

pub fn char_at(c: &mut Chunk, line: u32) {
    let idx = c.add_import("ecma:string", "charAt");
    c.emit_call(idx, 2, line);
}

pub fn index_of(c: &mut Chunk, line: u32) {
    let idx = c.add_import("ecma:string", "indexOf");
    c.emit_call(idx, 2, line);
}

pub fn last_index_of(c: &mut Chunk, line: u32) {
    let idx = c.add_import("ecma:string", "lastIndexOf");
    c.emit_call(idx, 2, line);
}

pub fn to_upper_case(c: &mut Chunk, line: u32) {
    let idx = c.add_import("ecma:string", "toUpperCase");
    c.emit_call(idx, 1, line);
}

pub fn to_lower_case(c: &mut Chunk, line: u32) {
    let idx = c.add_import("ecma:string", "toLowerCase");
    c.emit_call(idx, 1, line);
}

pub fn trim(c: &mut Chunk, line: u32) {
    let idx = c.add_import("ecma:string", "trim");
    c.emit_call(idx, 1, line);
}

pub fn trim_start(c: &mut Chunk, line: u32) {
    let idx = c.add_import("ecma:string", "trimStart");
    c.emit_call(idx, 1, line);
}

pub fn trim_end(c: &mut Chunk, line: u32) {
    let idx = c.add_import("ecma:string", "trimEnd");
    c.emit_call(idx, 1, line);
}

pub fn starts_with(c: &mut Chunk, line: u32) {
    let idx = c.add_import("ecma:string", "startsWith");
    c.emit_call(idx, 2, line);
}

pub fn ends_with(c: &mut Chunk, line: u32) {
    let idx = c.add_import("ecma:string", "endsWith");
    c.emit_call(idx, 2, line);
}

pub fn includes(c: &mut Chunk, line: u32) {
    let idx = c.add_import("ecma:string", "includes");
    c.emit_call(idx, 2, line);
}

pub fn replace(c: &mut Chunk, line: u32) {
    let idx = c.add_import("ecma:string", "replace");
    c.emit_call(idx, 3, line);
}

pub fn repeat(c: &mut Chunk, line: u32) {
    let idx = c.add_import("ecma:string", "repeat");
    c.emit_call(idx, 2, line);
}

pub fn pad_start(c: &mut Chunk, line: u32) {
    let idx = c.add_import("ecma:string", "padStart");
    c.emit_call(idx, 3, line);
}

pub fn pad_end(c: &mut Chunk, line: u32) {
    let idx = c.add_import("ecma:string", "padEnd");
    c.emit_call(idx, 3, line);
}

pub fn slice(c: &mut Chunk, line: u32) {
    let idx = c.add_import("ecma:string", "slice");
    c.emit_call(idx, 3, line);
}

/// Inline string reverse: split("") → array.reverse() → join("")
pub fn reverse(c: &mut Chunk, line: u32) {
    c.emit_string_const("", line);
    let split = c.add_import("ecma:string", "split");
    c.emit_call(split, 2, line);
    let rev = c.add_import("ecma:array", "reverse");
    c.emit_call(rev, 1, line);
    c.emit_string_const("", line);
    let join = c.add_import("ecma:array", "join");
    c.emit_call(join, 2, line);
}
