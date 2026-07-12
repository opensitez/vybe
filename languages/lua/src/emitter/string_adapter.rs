//! Lua string pattern adapter — Lua patterns → JS regex, then ecma:regexp.*
//!
//! Lua uses its own pattern syntax (%d, %a, %s, …) that is incompatible
//! with JS regex.  Each emit_lua_string_* function:
//!   1. Pops the args from the stack.
//!   2. Converts the Lua pattern to a JS regex string via a series of
//!      ecma:regexp:replaceAll calls (plain string replacements — no
//!      regex used for the conversion itself).
//!   3. Calls the appropriate ecma:regexp:* host fn with the converted
//!      pattern + the original string args.
//!
//! No new host fns; no polyfills.  Pure bytecode over ecma:regexp.*.

use std::sync::Arc;
use vybe_bytecode::opcode::Op;
use vybe_bytecode::{Chunk, Value};

// ── helpers ─────────────────────────────────────────────────────────

fn alloc_local(chunk: &mut Chunk) -> u16 {
    chunk.alloc_scratch(1)
}

fn lset(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn lget(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn push_str(chunk: &mut Chunk, s: &str, line: u32) {
    chunk.emit_string_const(s, line);
}

fn push_null(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::NULL, line);
}

/// Call a host import on `chunks[current]` with `argc` args already on stack.
fn call_import(
    chunks: &mut Vec<Chunk>,
    current: usize,
    module: &str,
    name: &str,
    argc: u8,
    line: u32,
) {
    let idx = chunks[current].add_import(module, name);
    chunks[current].emit_call(idx, argc, line);
}

// ── Lua pattern → JS regex conversion ───────────────────────────────
//
// Replaces each Lua character class escape with its JS regex equivalent.
// The conversion is done by a chain of ecma:regexp:replaceAll calls, each
// treating the pattern string as a *plain string* (no regex magic in the
// search — we pass a JS regex with the literal % sign escaped as \\%).
//
// Order matters: we do multi-char replacements before single-char ones
// to avoid double-substitution (e.g. %d before %D is fine because %D is
// handled separately, but doing %d → [0-9] could clobber if we're not
// careful).  Since we use replaceAll with literal patterns, there is no
// ambiguity.
//
// Stack on entry:  [lua_pattern: string]
// Stack on exit:   [js_regex: string]

// Each substitution: replace `lua_pat` → `js_repl` in the string on top
// of the stack.  replaceAll(str, search, replacement)  →  new string.
fn emit_replace_literal(
    chunks: &mut Vec<Chunk>,
    current: usize,
    lua_pat: &str,
    js_repl: &str,
    line: u32,
) {
    // Stack: [current_pattern]
    // We need: replaceAll(str, lua_pat_as_regex, js_repl)
    // Pass lua_pat wrapped in /.../ so ecma:regexp:replaceAll treats it as regex.
    // Escape special regex chars in lua_pat: only '%' matters here.
    // We use a JS regex /\%d/g style — wrap in /pattern/g.
    let escaped = escape_for_js_regex_pattern(lua_pat);
    let regex_str = format!("/{}/g", escaped);

    push_str(&mut chunks[current], &regex_str, line);
    push_str(&mut chunks[current], js_repl, line);
    call_import(chunks, current, "ecma:regexp", "replaceAll", 3, line);
}

/// Escape special JS regex metacharacters in a Lua pattern literal
/// so it can be wrapped in /.../ for a JS regex search.
fn escape_for_js_regex_pattern(s: &str) -> String {
    let mut out = String::with_capacity(s.len() * 2);
    for c in s.chars() {
        match c {
            '.' | '*' | '+' | '?' | '(' | ')' | '[' | ']' | '{' | '}' | '|' | '^' | '$' | '\\'
            | '/' => {
                out.push('\\');
                out.push(c);
            }
            '%' => {
                // In JS regex, % is not special, but we escape it for clarity.
                out.push('%');
            }
            _ => out.push(c),
        }
    }
    out
}

/// Emit bytecode that converts a Lua pattern (top of stack) to a JS regex string.
/// Stack in: [lua_pattern]  Stack out: [js_regex_string]
fn emit_lua_pattern_to_js_regex(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    // 1. Quantifier '-' replacements FIRST, before we introduce any hyphens by replacing classes.
    // Replace non-greedy class matches first, e.g. %d- -> [0-9]*?
    let quantifier_substitutions: &[(&str, &str)] = &[
        // Non-greedy class quantifiers
        ("%d-", "[0-9]*?"),
        ("%D-", "[^0-9]*?"),
        ("%a-", "[a-zA-Z]*?"),
        ("%A-", "[^a-zA-Z]*?"),
        ("%l-", "[a-z]*?"),
        ("%L-", "[^a-z]*?"),
        ("%u-", "[A-Z]*?"),
        ("%U-", "[^A-Z]*?"),
        ("%s-", "[\\t\\n\\r\\f\\v ]*?"),
        ("%S-", "[^\\t\\n\\r\\f\\v ]*?"),
        ("%w-", "[a-zA-Z0-9]*?"),
        ("%W-", "[^a-zA-Z0-9]*?"),
        ("%x-", "[0-9a-fA-F]*?"),
        ("%X-", "[^0-9a-fA-F]*?"),
        ("%p-", "[!-/:-@\\[-`{-~]*?"),
        ("%P-", "[^!-/:-@\\[-`{-~]*?"),
        ("%c-", "[\\x00-\\x1f\\x7f]*?"),
        ("%C-", "[^\\x00-\\x1f\\x7f]*?"),
        ("%g-", "[!-~]*?"),
        ("%G-", "[^!-~]*?"),
        // Non-greedy dot quantifier
        (".-", ".*?"),
        // Non-greedy set quantifier (end of set ] followed by -)
        ("]-", "]*?"),
    ];

    for &(from, to) in quantifier_substitutions {
        emit_replace_literal(chunks, current, from, to, line);
    }

    // 2. Character class replacements (without quantifiers, since those were handled above).
    let substitutions: &[(&str, &str)] = &[
        ("%d", "[0-9]"),
        ("%D", "[^0-9]"),
        ("%a", "[a-zA-Z]"),
        ("%A", "[^a-zA-Z]"),
        ("%l", "[a-z]"),
        ("%L", "[^a-z]"),
        ("%u", "[A-Z]"),
        ("%U", "[^A-Z]"),
        ("%s", "[\\t\\n\\r\\f\\v ]"),
        ("%S", "[^\\t\\n\\r\\f\\v ]"),
        ("%w", "[a-zA-Z0-9]"),
        ("%W", "[^a-zA-Z0-9]"),
        ("%x", "[0-9a-fA-F]"),
        ("%X", "[^0-9a-fA-F]"),
        ("%p", "[!-/:-@\\[-`{-~]"),
        ("%P", "[^!-/:-@\\[-`{-~]"),
        ("%c", "[\\x00-\\x1f\\x7f]"),
        ("%C", "[^\\x00-\\x1f\\x7f]"),
        ("%g", "[!-~]"),
        ("%G", "[^!-~]"),
    ];

    for &(from, to) in substitutions {
        emit_replace_literal(chunks, current, from, to, line);
    }

    // 3. Escape sequences and literal punctuation replacements.
    // Note: %- is the Lua escape for a literal hyphen. We replace it with \-
    // to match in JS regex.
    let punct_escapes: &[(&str, &str)] = &[
        ("%%", "%"),
        ("%.", "\\."),
        ("%+", "\\+"),
        ("%-", "\\-"),
        ("%*", "\\*"),
        ("%?", "\\?"),
        ("%(", "\\("),
        ("%)", "\\)"),
        ("%[", "\\["),
        ("%]", "\\]"),
        ("%{", "\\{"),
        ("%}", "\\}"),
        ("%^", "\\^"),
        ("%$", "\\$"),
        ("%|", "\\|"),
        ("%/", "\\/"),
        ("%\\", "\\\\"),
        ("%\"", "\\\""),
        ("%'", "\\'"),
        ("%#", "#"),
        ("%@", "@"),
        ("%!", "!"),
        ("%&", "&"),
        ("%,", ","),
        ("%;", ";"),
        ("%:", ":"),
        ("%=", "="),
        ("%<", "<"),
        ("%>", ">"),
        ("%~", "~"),
        ("%`", "`"),
    ];

    for &(from, to) in punct_escapes {
        emit_replace_literal(chunks, current, from, to, line);
    }
}

// ── string.match ────────────────────────────────────────────────────
//
// Lua: string.match(s, pat [, init])
// JS:  ecma:regexp.match(s, js_regex)  — returns match array or null
//
// If there are captures (groups) in the pattern, Lua returns the capture
// values; otherwise it returns the whole match.
// ecma:regexp.match returns an Array: [full_match, cap1, cap2, ...]
// So if captures exist (array length > 1), return cap1.
// Otherwise return full_match (index 0).
//
// Stack on entry: [..., s, pat] (argc = 2) or [..., s, pat, init] (argc = 3)
// Stack on exit:  [..., result | null]

pub fn emit_lua_string_match(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    // Alloc locals
    let (s_slot, pat_slot, init_slot, js_pat_slot, result_slot) = {
        let c = &mut chunks[current];
        (
            alloc_local(c),
            alloc_local(c),
            alloc_local(c),
            alloc_local(c),
            alloc_local(c),
        )
    };

    // Pop args: stack is [s, pat] or [s, pat, init] (last pushed = top)
    {
        let c = &mut chunks[current];
        if argc >= 3 {
            lset(c, init_slot, line);
        } else {
            push_null(c, line);
            lset(c, init_slot, line);
        }
        lset(c, pat_slot, line);
        lset(c, s_slot, line);
    }

    // Convert Lua pattern to JS regex
    lget(&mut chunks[current], pat_slot, line);
    emit_lua_pattern_to_js_regex(chunks, current, line);
    lset(&mut chunks[current], js_pat_slot, line);

    // If init is set, apply it as a slice offset to s before matching
    // For now: apply substring if init is not null
    // TODO: full init support — for MVP, pass s directly
    lget(&mut chunks[current], s_slot, line);
    lget(&mut chunks[current], js_pat_slot, line);
    call_import(chunks, current, "ecma:regexp", "match", 2, line);
    lset(&mut chunks[current], result_slot, line);

    // If result is null → push null
    // Else if result.length == 1 → push result[0] (no captures)
    // Else → push result[1] (first capture)
    {
        let c = &mut chunks[current];
        lget(c, result_slot, line);
        c.emit_op(Op::REF_IS_NULL, line);
        c.emit_if(line);
        push_null(c, line);
        c.emit_else(line);

        // result is not null — check number of captures
        lget(c, result_slot, line);
        c.emit_op(Op::ARRAY_LENGTH, line);
        c.emit_f64_const(1.0, line);
        c.emit_op(Op::F64_GT, line);
        c.emit_if(line);
        // Has captures — return result[1] (first capture group)
        lget(c, result_slot, line);
        c.emit_f64_const(1.0, line);
        c.emit_op(Op::ARRAY_GET, line);
        c.emit_else(line);
        // No captures — return result[0] (full match)
        lget(c, result_slot, line);
        c.emit_f64_const(0.0, line);
        c.emit_op(Op::ARRAY_GET, line);
        c.emit_end(line);

        c.emit_end(line);
    }
}

// ── string.find ─────────────────────────────────────────────────────
//
// Lua: string.find(s, pat [, init [, plain]])
// Returns: start, end (1-based) [, cap1, cap2, ...] or nil
//
// ecma:regexp.exec(regex_obj, str) returns Array with .index
// For MVP: returns start+1, end+1 (1-based) using match.index + match[0].length

pub fn emit_lua_string_find(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let (s_slot, pat_slot, _init_slot, _plain_slot, js_pat_slot, result_slot) = {
        let c = &mut chunks[current];
        (
            alloc_local(c),
            alloc_local(c),
            alloc_local(c),
            alloc_local(c),
            alloc_local(c),
            alloc_local(c),
        )
    };

    {
        let c = &mut chunks[current];
        if argc >= 4 {
            lset(c, _plain_slot, line);
        } else {
            push_null(c, line);
            lset(c, _plain_slot, line);
        }
        if argc >= 3 {
            lset(c, _init_slot, line);
        } else {
            push_null(c, line);
            lset(c, _init_slot, line);
        }
        lset(c, pat_slot, line);
        lset(c, s_slot, line);
    }

    // Convert pattern
    lget(&mut chunks[current], pat_slot, line);
    emit_lua_pattern_to_js_regex(chunks, current, line);
    lset(&mut chunks[current], js_pat_slot, line);

    // Build regex obj
    lget(&mut chunks[current], js_pat_slot, line);
    call_import(chunks, current, "ecma:regexp", "new", 1, line);
    // exec(regex_obj, str)
    lget(&mut chunks[current], s_slot, line);
    call_import(chunks, current, "ecma:regexp", "exec", 2, line);
    lset(&mut chunks[current], result_slot, line);

    // if result == null → push null
    {
        let c = &mut chunks[current];
        lget(c, result_slot, line);
        c.emit_op(Op::REF_IS_NULL, line);
        c.emit_if(line);
        push_null(c, line);
        c.emit_else(line);
        // start = result.index + 1  (Lua 1-based)
        let index_k = c.add_constant(Value::String(Arc::from("index")));
        lget(c, result_slot, line);
        c.emit_op_u16(Op::STRUCT_GET, index_k, line);
        c.emit_f64_const(1.0, line);
        c.emit_op(Op::F64_ADD, line);
        // end = start - 1 + len(result[0])  = result.index + len(result[0])
        lget(c, result_slot, line);
        let idx_k = c.add_constant(Value::String(Arc::from("index")));
        c.emit_op_u16(Op::STRUCT_GET, idx_k, line);
        lget(c, result_slot, line);
        c.emit_f64_const(0.0, line);
        c.emit_op(Op::ARRAY_GET, line);
        let len_idx = c.add_import("wasm:js-string", "length");
        c.emit_call(len_idx, 1, line);
        c.emit_op(Op::F64_ADD, line);
        // Stack: [start, end]  — Lua returns them as a multi-value
        // For now return as array so the multi-assign desugaring works
        c.emit_op_u16(Op::ARRAY_NEW_FIXED, 2, line);
        c.emit_end(line);
    }
}

// ── string.gsub ─────────────────────────────────────────────────────
//
// Lua: string.gsub(s, pat, repl [, n])
// repl can be: string, table, or function
// Returns: new_str, count
//
// For MVP: string replacement using ecma:regexp:replace (with g flag for all).
// The count return is approximated.

pub fn emit_lua_string_gsub(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let (s_slot, pat_slot, repl_slot, _n_slot, js_pat_slot) = {
        let c = &mut chunks[current];
        (
            alloc_local(c),
            alloc_local(c),
            alloc_local(c),
            alloc_local(c),
            alloc_local(c),
        )
    };

    {
        let c = &mut chunks[current];
        if argc >= 4 {
            lset(c, _n_slot, line);
        } else {
            push_null(c, line);
            lset(c, _n_slot, line);
        }
        lset(c, repl_slot, line);
        lset(c, pat_slot, line);
        lset(c, s_slot, line);
    }

    // Convert pattern — append 'g' flag for global replace
    lget(&mut chunks[current], pat_slot, line);
    emit_lua_pattern_to_js_regex(chunks, current, line);
    // Wrap as /pattern/g  — prepend '/' and append '/g'
    {
        let c = &mut chunks[current];
        let js_pat_tmp = alloc_local(c);
        lset(c, js_pat_tmp, line);
        // Build "/" + pattern + "/g"
        push_str(c, "/", line);
        lget(c, js_pat_tmp, line);
        vybe_emitter::ops::emit_dyn_add(c, line);
        push_str(c, "/g", line);
        vybe_emitter::ops::emit_dyn_add(c, line);
        lset(c, js_pat_slot, line);
    }

    // ecma:regexp:replace(str, pattern, replacement)
    lget(&mut chunks[current], s_slot, line);
    lget(&mut chunks[current], js_pat_slot, line);
    lget(&mut chunks[current], repl_slot, line);
    call_import(chunks, current, "ecma:regexp", "replace", 3, line);
    // Lua gsub returns (new_str, count) but our multi-return is limited.
    // Return just the string for now (count = 0 placeholder).
    // TODO: track actual replacement count.
}

// ── __lua_gmatch_match_all ──────────────────────────────────────────
//
// Strategy: use ecma:regexp:matchAll to get all matches up front as an
// array and leave it on the stack.

pub fn emit_lua_string_gmatch_match_all(
    chunks: &mut Vec<Chunk>,
    current: usize,
    _argc: u8,
    line: u32,
) {
    let (s_slot, pat_slot, js_pat_slot, matches_slot) = {
        let c = &mut chunks[current];
        (
            alloc_local(c),
            alloc_local(c),
            alloc_local(c),
            alloc_local(c),
        )
    };

    {
        let c = &mut chunks[current];
        lset(c, pat_slot, line);
        lset(c, s_slot, line);
    }

    // Convert pattern with 'g' flag
    lget(&mut chunks[current], pat_slot, line);
    emit_lua_pattern_to_js_regex(chunks, current, line);
    {
        let c = &mut chunks[current];
        let tmp = alloc_local(c);
        lset(c, tmp, line);
        push_str(c, "/", line);
        lget(c, tmp, line);
        vybe_emitter::ops::emit_dyn_add(c, line);
        push_str(c, "/g", line);
        vybe_emitter::ops::emit_dyn_add(c, line);
        lset(c, js_pat_slot, line);
    }

    // ecma:regexp:matchAll(str, pattern) → array of match arrays
    lget(&mut chunks[current], s_slot, line);
    lget(&mut chunks[current], js_pat_slot, line);
    call_import(chunks, current, "ecma:regexp", "matchAll", 2, line);
    lset(&mut chunks[current], matches_slot, line);

    // Leave matches array on the stack
    lget(&mut chunks[current], matches_slot, line);
}
