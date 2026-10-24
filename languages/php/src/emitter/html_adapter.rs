//! php `htmlspecialchars` and `preg_quote` as guest code over `ecma:*`
//! string and RegExp builtins. Patterns are compiled once per request and
//! kept in guest globals; the escaping itself is one helper function per
//! module.
use vybe_compiler::primitives::{globals, ops};
use vybe_runtime::{Chunk, opcode::Op};

use crate::html_entity_names::{HTML4_ENTITY_NAMES, HTML5_ENTITY_NAMES};

const ENT_QUOTES_SINGLE: i32 = 1;
const ENT_COMPAT: i32 = 2;
const ENT_DISALLOWED: i32 = 128;
const DOCUMENT_MASK: i32 = 48;
const XML1: i32 = 16;
const XHTML: i32 = 32;
const HTML5: i32 = 48;
/// ENT_QUOTES | ENT_SUBSTITUTE | ENT_HTML401.
const DEFAULT_FLAGS: f64 = 11.0;

const SPECIALCHARS_HELPER: &str = "__php_htmlspecialchars";

fn import(chunk: &mut Chunk, module: &str, name: &str, argc: u8, line: u32) {
    let index = chunk.add_import(module, name);
    chunk.emit_call(index, argc, line);
}

fn get(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn set(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

/// Push the RegExp `pattern`/`flags`, compiled on first use into `global`.
fn emit_cached_regexp(chunk: &mut Chunk, global: &str, pattern: &str, flags: &str, line: u32) {
    let regexp = chunk.alloc_scratch(1);
    globals::emit_read(chunk, global, line);
    chunk.emit_op_u16(Op::LOCAL_TEE, regexp, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    get(chunk, regexp, line);
    import(chunk, "wasm:js-undefined", "test", 1, line);
    chunk.emit_op(Op::I32_OR, line);
    chunk.emit_if(line);
    chunk.emit_string_const(pattern, line);
    chunk.emit_string_const(flags, line);
    import(chunk, "ecma:regexp", "new", 2, line);
    chunk.emit_op_u16(Op::LOCAL_TEE, regexp, line);
    globals::emit_write(chunk, global, line);
    chunk.emit_end(line);
    get(chunk, regexp, line);
}

/// `text = text.replaceAll(search, replacement)`, literally.
fn emit_replace_literal(chunk: &mut Chunk, text: u16, search: &str, replacement: &str, line: u32) {
    get(chunk, text, line);
    chunk.emit_string_const(search, line);
    chunk.emit_string_const(replacement, line);
    import(chunk, "ecma:string", "replaceAll", 3, line);
    set(chunk, text, line);
}

/// `text = text.replaceAll(<cached RegExp>, replacement)`.
fn emit_replace_pattern(
    chunk: &mut Chunk, text: u16, global: &str, pattern: &str, flags: &str, replacement: &str,
    line: u32,
) {
    get(chunk, text, line);
    emit_cached_regexp(chunk, global, pattern, flags, line);
    chunk.emit_string_const(replacement, line);
    import(chunk, "ecma:regexp", "replaceAll", 3, line);
    set(chunk, text, line);
}

/// An `&` that does not start an entity reference htmlspecialchars keeps when
/// `double_encode` is off: a named entity of the document type, or a decimal
/// or hex reference to a code point no greater than U+10FFFF.
fn bare_ampersand_pattern(names: &[&str]) -> String {
    let decimal = "#0*(?:[0-9]{1,6}|10[0-9]{5}|110[0-9]{4}|111[0-3][0-9]{3}|1114(?:0[0-9]{2}|10[0-9]|11[01]))";
    let hex = "#[xX]0*(?:[0-9a-fA-F]{1,5}|10[0-9a-fA-F]{4})";
    format!("&(?!(?:{decimal}|{hex}|{});)", names.join("|"))
}

/// Characters a document type disallows, replaced under ENT_DISALLOWED.
fn disallowed_pattern(document: i32) -> &'static str {
    match document {
        0 => r"[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F-\u009F]",
        HTML5 => concat!(
            r"[\u0000-\u0008\u000B\u000E-\u001F\u007F-\u009F﷐-﷯￾￿",
            r"\u{1FFFE}\u{1FFFF}\u{2FFFE}\u{2FFFF}\u{3FFFE}\u{3FFFF}\u{4FFFE}\u{4FFFF}",
            r"\u{5FFFE}\u{5FFFF}\u{6FFFE}\u{6FFFF}\u{7FFFE}\u{7FFFF}\u{8FFFE}\u{8FFFF}",
            r"\u{9FFFE}\u{9FFFF}\u{AFFFE}\u{AFFFF}\u{BFFFE}\u{BFFFF}\u{CFFFE}\u{CFFFF}",
            r"\u{DFFFE}\u{DFFFF}\u{EFFFE}\u{EFFFF}\u{FFFFE}\u{FFFFF}\u{10FFFE}\u{10FFFF}]",
        ),
        _ => r"[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]",
    }
}

/// For each document type in order, emit `if (document == type) { body }
/// else {`, then the default body, then the closing ends.
fn emit_per_document(
    chunk: &mut Chunk, document: u16, line: u32, mut body: impl FnMut(&mut Chunk, i32),
) {
    for value in [XML1, XHTML, HTML5] {
        get(chunk, document, line);
        chunk.emit_i32_const(value, line);
        chunk.emit_op(Op::I32_EQ, line);
        chunk.emit_if(line);
        body(chunk, value);
        chunk.emit_else(line);
    }
    body(chunk, 0);
    for _ in 0..3 {
        chunk.emit_end(line);
    }
}

/// `__php_htmlspecialchars(text, flags, double_encode)`.
fn ensure_specialchars_helper(chunks: &mut Vec<Chunk>) -> usize {
    if let Some(index) = chunks.iter().position(|chunk| chunk.name == SPECIALCHARS_HELPER) {
        return index;
    }
    let index = chunks.len();
    let mut helper = Chunk::new(SPECIALCHARS_HELPER);
    helper.arity = 3;
    helper.local_count = 3;
    chunks.push(helper);
    let line = 0;
    let c = &mut chunks[index];
    let (text, flags_value, double_value) = (0u16, 1u16, 2u16);
    let flags = c.alloc_scratch(1);
    let document = c.alloc_scratch(1);

    get(c, flags_value, line);
    import(c, "ecma:number", "Number", 1, line);
    import(c, "wasm:js-number", "toF64", 1, line);
    c.emit_op(Op::I32_TRUNC_SAT_F64_S, line);
    set(c, flags, line);
    get(c, flags, line);
    c.emit_i32_const(DOCUMENT_MASK, line);
    c.emit_op(Op::I32_AND, line);
    set(c, document, line);

    // Nothing to escape and no disallowed-character policy: unchanged.
    get(c, flags, line);
    c.emit_i32_const(ENT_DISALLOWED, line);
    c.emit_op(Op::I32_AND, line);
    c.emit_op(Op::I32_EQZ, line);
    c.emit_if(line);
    emit_cached_regexp(c, "__php_html_special", "[&<>\"']", "", line);
    get(c, text, line);
    import(c, "ecma:regexp", "test", 2, line);
    import(c, "wasm:js-boolean", "cast", 1, line);
    c.emit_op(Op::I32_EQZ, line);
    c.emit_if(line);
    get(c, text, line);
    c.emit_op(Op::RETURN, line);
    c.emit_end(line);
    c.emit_end(line);

    // Disallowed characters first: none of them is one of the five, and none
    // can occur inside an entity reference.
    get(c, flags, line);
    c.emit_i32_const(ENT_DISALLOWED, line);
    c.emit_op(Op::I32_AND, line);
    c.emit_if(line);
    emit_per_document(c, document, line, |c, document| {
        let global = format!("__php_html_disallowed_{document}");
        emit_replace_pattern(c, text, &global, disallowed_pattern(document), "gu", "\u{FFFD}", line);
    });
    c.emit_end(line);

    // `&` before the other replacements introduce their own.
    get(c, double_value, line);
    ops::emit_dyn_to_bool(c, line);
    c.emit_if(line);
    emit_replace_literal(c, text, "&", "&amp;", line);
    c.emit_else(line);
    emit_per_document(c, document, line, |c, document| {
        let global = format!("__php_html_bare_ampersand_{document}");
        let pattern = match document {
            XML1 => bare_ampersand_pattern(&["amp", "lt", "gt", "quot", "apos"]),
            HTML5 => bare_ampersand_pattern(HTML5_ENTITY_NAMES),
            XHTML => {
                let mut names = vec!["apos"];
                names.extend_from_slice(HTML4_ENTITY_NAMES);
                bare_ampersand_pattern(&names)
            }
            _ => bare_ampersand_pattern(HTML4_ENTITY_NAMES),
        };
        emit_replace_pattern(c, text, &global, &pattern, "g", "&amp;", line);
    });
    c.emit_end(line);

    emit_replace_literal(c, text, "<", "&lt;", line);
    emit_replace_literal(c, text, ">", "&gt;", line);
    get(c, flags, line);
    c.emit_i32_const(ENT_COMPAT, line);
    c.emit_op(Op::I32_AND, line);
    c.emit_if(line);
    emit_replace_literal(c, text, "\"", "&quot;", line);
    c.emit_end(line);
    get(c, flags, line);
    c.emit_i32_const(ENT_QUOTES_SINGLE, line);
    c.emit_op(Op::I32_AND, line);
    c.emit_if(line);
    get(c, document, line);
    c.emit_op(Op::I32_EQZ, line);
    c.emit_if(line);
    emit_replace_literal(c, text, "'", "&#039;", line);
    c.emit_else(line);
    emit_replace_literal(c, text, "'", "&apos;", line);
    c.emit_end(line);
    c.emit_end(line);
    get(c, text, line);
    c.emit_op(Op::RETURN, line);
    index
}

/// php `htmlspecialchars($string, $flags = ENT_QUOTES | ENT_SUBSTITUTE |
/// ENT_HTML401, $encoding = null, $double_encode = true)`.
/// Stack: `[string, flags?, encoding?, double_encode?] -> [string]`, with the
/// string already converted.
pub fn emit_htmlspecialchars(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let c = &mut chunks[current];
    let double_encode = c.alloc_scratch(1);
    let flags = c.alloc_scratch(1);
    let text = c.alloc_scratch(1);
    if argc >= 4 {
        set(c, double_encode, line);
    } else {
        c.emit_bool_const(true, line);
        set(c, double_encode, line);
    }
    if argc >= 3 {
        c.emit_op(Op::DROP, line);
    }
    if argc >= 2 {
        set(c, flags, line);
    } else {
        c.emit_f64_const(DEFAULT_FLAGS, line);
        set(c, flags, line);
    }
    set(c, text, line);
    let helper = ensure_specialchars_helper(chunks);
    let c = &mut chunks[current];
    get(c, text, line);
    get(c, flags, line);
    get(c, double_encode, line);
    if vybe_compiler::primitives::functions::try_emit_direct_helper_call(chunks, current, helper, 3, line) {
        return;
    }
    let c = &mut chunks[current];
    let arguments = c.alloc_scratch(3);
    for offset in (0..3).rev() {
        set(c, arguments + offset, line);
    }
    c.emit_op_u16(Op::REF_FUNC, helper as u16, line);
    c.emit(0u8, line);
    for offset in 0..3 {
        get(c, arguments + offset, line);
    }
    vybe_compiler::primitives::callable::emit_direct_invoke_chunk(c, 3, line);
}

/// php `preg_quote($str, $delimiter = null)`: a backslash before each of
/// `. \ + * ? [ ^ ] $ ( ) { } = ! < > | : - #` and before the delimiter's
/// character, and `\000` for NUL. Stack: `[string, delimiter?] -> [string]`,
/// both already strings.
pub fn emit_preg_quote(chunk: &mut Chunk, argc: u8, line: u32) {
    let text = chunk.alloc_scratch(1);
    let delimiter = chunk.alloc_scratch(1);
    if argc >= 2 {
        set(chunk, delimiter, line);
    }
    set(chunk, text, line);
    emit_replace_pattern(
        chunk, text, "__php_preg_quote_special", r"[.\\+*?\[^\]$(){}=!<>|:\-#]", "g", "\\$&", line,
    );
    if argc >= 2 {
        // The delimiter's first character, when it is not already escaped.
        let first = chunk.alloc_scratch(1);
        get(chunk, delimiter, line);
        chunk.emit_i32_const(0, line);
        chunk.emit_i32_const(1, line);
        import(chunk, "wasm:js-string", "substring", 3, line);
        set(chunk, first, line);
        emit_cached_regexp(chunk, "__php_preg_quote_plain", r"^[^.\\+*?\[^\]$(){}=!<>|:\-#\u0000]$", "", line);
        get(chunk, first, line);
        import(chunk, "ecma:regexp", "test", 2, line);
        import(chunk, "wasm:js-boolean", "cast", 1, line);
        chunk.emit_if(line);
        get(chunk, text, line);
        get(chunk, first, line);
        chunk.emit_string_const("\\", line);
        get(chunk, first, line);
        import(chunk, "wasm:js-string", "concat", 2, line);
        import(chunk, "ecma:string", "replaceAll", 3, line);
        set(chunk, text, line);
        chunk.emit_end(line);
    }
    emit_replace_literal(chunk, text, "\0", "\\000", line);
    get(chunk, text, line);
}
