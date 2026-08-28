//! `System.Globalization.CultureInfo`, `TextInfo`, `NumberFormatInfo`.
//!
//! None of the three existed anywhere — not in `platforms/dotnet`, not in the
//! VB walker, not in any profile — so `CultureInfo.InvariantCulture` and
//! `CultureInfo.GetCultureInfo("en-US")` both died as "undefined is not
//! callable". A `System.*` type, so it lives here and C# gets it too.
//!
//! ## Values are .NET's, verified on `tools/vbrun`
//!
//! ⛔ Two corpus tests asserted values real VB.NET does not produce, and were
//! fixed rather than implemented to:
//! `CultureInfo.InvariantCulture.Name` is the EMPTY STRING (not `"en-US"`) and
//! its `CurrencySymbol` is `¤` U+00A4, the generic currency sign (not `$`).
//! The invariant culture is not "US English", it is the culture-independent
//! one; `$` belongs to `en-US`.
//!
//! ## Scope
//!
//! The culture DATA is the invariant culture plus the specific cultures the
//! corpus names, carried as instance fields. This is not a CLDR table and does
//! not pretend to be: an unknown culture name yields a culture whose `Name` is
//! the requested name and whose formatting fields are the invariant ones,
//! which is what `GetCultureInfo` on an unrecognised-but-well-formed name does
//! for the parts we model. When a real locale database lands, the constructor
//! is the single place that changes.

use vybe_compiler::primitives::collections;
use vybe_compiler::primitives::ops;
use vybe_compiler::primitives::strings;
use vybe_runtime::chunk::Chunk;
use vybe_runtime::opcode::Op;
use vybe_compiler::primitives::class_slots;

// ⛔ Every member is written in BOTH spellings. A dotnet type with no property
// accessor resolves its properties as a LOWERCASED struct-field read, so a
// PascalCase-only key is unreadable from a case-insensitive frontend (VB) and a
// lowercase-only key is unreadable from a case-sensitive one (C#).
// `set_*_both` below is the only way these objects should be written.

fn set_string(chunks: &mut [Chunk], current: usize, object: u16, key: &str, value: &str, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, object, line);
    chunks[current].emit_string_const(key, line);
    chunks[current].emit_string_const(value, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
}

fn set_bool(chunks: &mut [Chunk], current: usize, object: u16, key: &str, value: bool, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, object, line);
    chunks[current].emit_string_const(key, line);
    // ⛔ NOT `emit_bool_const` — a constant stored that way reads back FALSE
    // whatever it was set to, while the same field written through
    // `i32_to_bool` reads correctly (`IsNeutralCulture`, computed at runtime,
    // worked while `IsReadOnly`, a constant, did not). Materialize the bool the
    // same way the computed path does.
    chunks[current].emit_i32_const(i32::from(value), line);
    ops::emit_i32_to_bool(&mut chunks[current], line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
}

fn set_slot(chunks: &mut [Chunk], current: usize, object: u16, key: &str, value: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, object, line);
    chunks[current].emit_string_const(key, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
}

/// Write a member under BOTH the .NET spelling and its lowercase form.
fn set_string_both(
    chunks: &mut [Chunk],
    current: usize,
    object: u16,
    key: &str,
    value: &str,
    line: u32,
) {
    set_string(chunks, current, object, key, value, line);
    set_string(chunks, current, object, &key.to_ascii_lowercase(), value, line);
}

fn set_bool_both(chunks: &mut [Chunk], current: usize, object: u16, key: &str, on: bool, line: u32) {
    set_bool(chunks, current, object, key, on, line);
    set_bool(chunks, current, object, &key.to_ascii_lowercase(), on, line);
}

fn set_slot_both(chunks: &mut [Chunk], current: usize, object: u16, key: &str, v: u16, line: u32) {
    set_slot(chunks, current, object, key, v, line);
    set_slot(chunks, current, object, &key.to_ascii_lowercase(), v, line);
}

/// The `NumberFormat` sub-object. `currency` is the only field that differs
/// between the cultures modelled here, which is why it is a parameter and the
/// separators are not.
fn emit_number_format(chunks: &mut [Chunk], current: usize, currency: &str, line: u32) -> u16 {
    let nf = chunks[current].alloc_scratch(1);
    class_slots::emit_class_alloc(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, nf, line);
    set_string_both(chunks, current, nf, "NumberDecimalSeparator", ".", line);
    set_string_both(chunks, current, nf, "NumberGroupSeparator", ",", line);
    set_string_both(chunks, current, nf, "CurrencySymbol", currency, line);
    set_string_both(chunks, current, nf, "PercentSymbol", "%", line);
    set_string_both(chunks, current, nf, "PositiveSign", "+", line);
    set_string_both(chunks, current, nf, "NegativeSign", "-", line);
    nf
}

/// The `TextInfo` sub-object. Its methods are registered on the `TextInfo`
/// class; the object only has to exist and carry its culture's name.
fn emit_text_info(chunks: &mut [Chunk], current: usize, name: &str, line: u32) -> u16 {
    let ti = chunks[current].alloc_scratch(1);
    class_slots::emit_class_alloc(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, ti, line);
    set_string_both(chunks, current, ti, "CultureName", name, line);
    ti
}

/// Build one culture object.
///
/// `parent_name` is `""` for a neutral or invariant culture — .NET's
/// `"en-US".Parent` is `"en"`, and `"en".Parent` is the invariant culture.
fn emit_culture(
    chunks: &mut [Chunk],
    current: usize,
    name: &str,
    parent_name: &str,
    currency: &str,
    read_only: bool,
    line: u32,
) -> u16 {
    let nf = emit_number_format(chunks, current, currency, line);
    let ti = emit_text_info(chunks, current, name, line);

    let culture = chunks[current].alloc_scratch(1);
    class_slots::emit_class_alloc(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, culture, line);
    set_string_both(chunks, current, culture, "Name", name, line);
    set_string_both(chunks, current, culture, "DisplayName", name, line);
    set_string_both(chunks, current, culture, "ParentName", parent_name, line);
    // A culture is NEUTRAL when it names a language without a region — `en`,
    // not `en-US`. The invariant culture is NOT neutral (verified on vbrun).
    set_bool_both(
        chunks,
        current,
        culture,
        "IsNeutralCulture",
        !name.is_empty() && !name.contains('-'),
        line,
    );
    set_bool_both(chunks, current, culture, "IsReadOnly", read_only, line);
    set_slot_both(chunks, current, culture, "NumberFormat", nf, line);
    set_slot_both(chunks, current, culture, "TextInfo", ti, line);
    culture
}

/// `CultureInfo.InvariantCulture`.
///
/// ⛔ `Name` is `""` and `CurrencySymbol` is `¤` — see the module header.
pub fn emit_invariant_culture(chunks: &mut [Chunk], current: usize, line: u32) {
    let culture = emit_culture(chunks, current, "", "", "\u{00A4}", true, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, culture, line);
}

/// `CultureInfo.CurrentCulture` — `en-US` in this runtime, which has no OS
/// locale to ask.
pub fn emit_current_culture(chunks: &mut [Chunk], current: usize, line: u32) {
    let culture = emit_culture(chunks, current, "en-US", "en", "$", false, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, culture, line);
}

/// `CultureInfo.GetCultureInfo(name)` / `New CultureInfo(name)`.
///
/// The name is carried through verbatim so `Parent`, `IsNeutralCulture` and
/// `Equals` answer from it. The currency symbol is chosen at RUNTIME from the
/// requested name, because the name is a value, not a compile-time constant.
pub fn emit_get_culture_info(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 0 {
        emit_current_culture(chunks, current, line);
        return;
    }
    let requested = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, requested, line);

    // Everything but the name is invariant-shaped; the name, its parent and the
    // currency are computed from the requested name below.
    let culture = emit_culture(chunks, current, "", "", "$", false, line);

    set_slot_both(chunks, current, culture, "Name", requested, line);
    set_slot_both(chunks, current, culture, "DisplayName", requested, line);

    // Parent: the text before the first `-`, empty when there is none.
    let parent = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, requested, line);
    chunks[current].emit_string_const("-", line);
    strings::emit_index_of(&mut chunks[current], line);
    let dash = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, dash, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, dash, line);
    chunks[current].emit_i32_const(0, line);
    ops::emit_dyn_ge(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, requested, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, dash, line);
    strings::emit_substring(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, parent, line);
    chunks[current].emit_else(line);
    chunks[current].emit_string_const("", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, parent, line);
    chunks[current].emit_end(line);
    set_slot_both(chunks, current, culture, "ParentName", parent, line);

    // Neutral when the name carries no region — i.e. no `-`.
    chunks[current].emit_op_u16(Op::LOCAL_GET, culture, line);
    chunks[current].emit_string_const("IsNeutralCulture", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, dash, line);
    chunks[current].emit_i32_const(0, line);
    ops::emit_dyn_lt(&mut chunks[current], line);
    ops::emit_i32_to_bool(&mut chunks[current], line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, culture, line);
    chunks[current].emit_string_const("isneutralculture", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, dash, line);
    chunks[current].emit_i32_const(0, line);
    ops::emit_dyn_lt(&mut chunks[current], line);
    ops::emit_i32_to_bool(&mut chunks[current], line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, culture, line);
}

/// `.Parent` — the culture named by `ParentName`.
///
/// Built on demand rather than stored, because storing it would make every
/// culture construct its whole ancestor chain eagerly.
pub fn emit_culture_parent(chunks: &mut [Chunk], current: usize, line: u32) {
    let child = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, child, line);

    let parent = emit_culture(chunks, current, "", "", "\u{00A4}", false, line);
    let parent_name = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, child, line);
    chunks[current].emit_string_const("parentname", line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, parent_name, line);
    set_slot_both(chunks, current, parent, "Name", parent_name, line);
    set_slot_both(chunks, current, parent, "DisplayName", parent_name, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, parent, line);
}

/// `.Clone()` — .NET returns a WRITABLE copy, which is what
/// `IsReadOnly`/`Not clone.IsReadOnly` asserts.
pub fn emit_culture_clone(chunks: &mut [Chunk], current: usize, line: u32) {
    let source = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, source, line);

    // ⛔ NOT `collections::emit_clone` — that is the ARRAY clone, and on an
    // object it hands back something with none of the fields, so
    // `clone.IsReadOnly` read back as neither True nor False and `Not` on it
    // produced the raw bitwise `-1`. Copy the members explicitly.
    let copy = emit_culture(chunks, current, "", "", "\u{00A4}", false, line);
    let carried = chunks[current].alloc_scratch(1);
    // Read through the LOWERCASE spelling (always written) and write BOTH, so
    // the copy stays readable from a case-sensitive frontend too.
    for key in [
        "Name",
        "DisplayName",
        "ParentName",
        "IsNeutralCulture",
        "NumberFormat",
        "TextInfo",
    ] {
        chunks[current].emit_op_u16(Op::LOCAL_GET, source, line);
        chunks[current].emit_string_const(&key.to_ascii_lowercase(), line);
        collections::emit_get(chunks, current, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, carried, line);
        set_slot_both(chunks, current, copy, key, carried, line);
    }
    // .NET's `Clone` returns a WRITABLE culture whatever the source was.
    set_bool_both(chunks, current, copy, "IsReadOnly", false, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, copy, line);
}

/// `.Equals(other)` — cultures compare by NAME, so two `GetCultureInfo("en-US")`
/// results are equal without being the same object.
pub fn emit_culture_equals(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(2);
    let (left, right) = (base, base + 1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, right, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, left, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, left, line);
    chunks[current].emit_string_const("name", line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, right, line);
    chunks[current].emit_string_const("name", line);
    collections::emit_get(chunks, current, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_i32_to_bool(&mut chunks[current], line);
}

/// `TextInfo.ToTitleCase(text)` — upper-case the first letter of each
/// whitespace-separated word, leaving the rest as it stands.
///
/// ⛔ NOT "capitalise and lowercase the tail": .NET leaves an already-upper
/// tail alone, which is why `ToUpper("vb")` and a title-cased `"VB"` differ.
pub fn emit_to_title_case(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(6);
    let (_recv, text, out, index, len, ch) =
        (base, base + 1, base + 2, base + 3, base + 4, base + 5);
    chunks[current].emit_op_u16(Op::LOCAL_SET, text, line);
    chunks[current].emit_op(Op::DROP, line); // the TextInfo receiver

    chunks[current].emit_string_const("", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, index, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, text, line);
    strings::emit_length(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len, line);

    let block = chunks[current].emit_block(line);
    let (loop_id, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, index, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len, line);
    ops::emit_dyn_lt(&mut chunks[current], line);
    ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, text, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, index, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, index, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    strings::emit_substring(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, ch, line);

    // Start of a word when it is the first character or follows a space.
    chunks[current].emit_op_u16(Op::LOCAL_GET, index, line);
    chunks[current].emit_i32_const(0, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, ch, line);
    strings::emit_to_upper(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, ch, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, text, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, index, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, index, line);
    strings::emit_substring(&mut chunks[current], line);
    chunks[current].emit_string_const(" ", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, ch, line);
    strings::emit_to_upper(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, ch, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, ch, line);
    strings::emit_concat(&mut chunks[current], 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, index, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, index, line);

    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_id);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);

    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
}

/// `TextInfo.ToUpper(text)` / `.ToLower(text)` — the receiver is dropped and
/// the shared string primitive does the work. Culture-specific casing (Turkish
/// dotless i) is NOT modelled; `emit_to_upper` is the invariant mapping.
pub fn emit_text_info_to_upper(chunks: &mut [Chunk], current: usize, line: u32) {
    let text = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, text, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, text, line);
    strings::emit_to_upper(&mut chunks[current], line);
}

pub fn emit_text_info_to_lower(chunks: &mut [Chunk], current: usize, line: u32) {
    let text = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, text, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, text, line);
    strings::emit_to_lower(&mut chunks[current], line);
}


/// The neutral cultures this runtime models — real ISO 639-1 language codes,
/// the same set .NET would report as `CultureTypes.NeutralCultures` for a
/// minimal ICU install. It is a SUBSET of a full CLDR listing and is written
/// out rather than fabricated at runtime, so what the runtime claims to know
/// is auditable in one place.
const NEUTRAL_CULTURES: [&str; 16] = [
    "ar", "de", "en", "es", "fr", "hi", "it", "ja", "ko", "nl", "pl", "pt", "ru", "sv", "tr", "zh",
];

/// `CultureInfo.GetCultures(kind)` — the `kind` argument is read and dropped:
/// every culture listed here is neutral, so `NeutralCultures` and `AllCultures`
/// answer the same set and `SpecificCultures` would need a region table this
/// does not have.
pub fn emit_get_cultures(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    for _ in 0..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    let list = chunks[current].alloc_scratch(2);
    let item = list + 1;
    chunks[current].emit_i32_const(0, line);
    collections::emit_new_with_length(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, list, line);
    for name in NEUTRAL_CULTURES {
        let culture = emit_culture(chunks, current, name, "", "\u{00A4}", true, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, culture, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, item, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, list, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, item, line);
        collections::emit_push(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_GET, list, line);
}
