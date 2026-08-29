//! PowerShell operator lowerings that need a RUNTIME type test.

use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

/// `@( … )` — the array subexpression operator. It guarantees an array and
/// flattens one level: `@(1..5)` is five elements, `@($arr)` is `$arr`'s
/// elements, and `@(7)` is one element.
///
/// Whether an operand flattens is a question about its RUNTIME value, so it
/// cannot be decided from the syntax: marking elements as spread drops a scalar
/// (nothing to iterate), and nesting them keeps a collection whole. `concat`
/// answers that for an array, a Set, a Map and a scalar alike.
///
/// ⛔**BUT NOT FOR A STRING.** `ecma:array:concat` deliberately serves TWO
/// callers — spec `Array.prototype.concat` AND spread-element compilation
/// (`[...s]`) — and its own comment says so, so it spreads a string into its
/// code points. That is right for `[...s]` and wrong for both `concat` and
/// `@( … )`: PowerShell's `@("abc")` is ONE element, and it was answering
/// THREE. Measured against `/usr/local/bin/pwsh` — `@("abc").Count` is 1 there
/// and was 3 here, so `@($text)` handed back the text's characters and every
/// `$s.Count` / `$s[0]` on it read a letter.
///
/// The string is therefore appended, never concatenated. Everything else keeps
/// going through `concat`, so an array still flattens and a Set/Map still
/// enumerates — testing `isArray` instead would have taken those two with it.
///
/// ⛔This is not an `ecma` bug to fix in `ecma`: that host fn has a second
/// caller which NEEDS the spreading. The wrong meaning was being asked for.
///
/// Stack: `[a0, …, aN-1]` → `[array]`.
pub fn emit_ensure_array(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let concat = chunk.add_import("ecma:array", "concat");
    let push = chunk.add_import("ecma:array", "push");
    let is_string = chunk.add_import("wasm:js-string", "test");

    // Arguments arrive with the LAST on top, so store them before rebuilding
    // the call in source order.
    let base = chunk.alloc_scratch(argc.max(1) as u16);
    for i in (0..argc as u16).rev() {
        chunk.emit_op_u16(Op::LOCAL_SET, base + i, line);
    }

    // `concat` returns a NEW array while `push` mutates in place, so the
    // accumulator lives in a local rather than on the stack — the two branches
    // cannot leave the same thing there.
    let acc = chunk.alloc_scratch(1);
    chunk.emit_array_new_fixed(0, 0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, acc, line);

    for i in 0..argc as u16 {
        chunk.emit_op_u16(Op::LOCAL_GET, base + i, line);
        chunk.emit_call(is_string, 1, line);
        chunk.emit_if(line);

        chunk.emit_op_u16(Op::LOCAL_GET, acc, line);
        chunk.emit_op_u16(Op::LOCAL_GET, base + i, line);
        chunk.emit_call(push, 2, line);
        chunk.emit_op(Op::DROP, line);

        chunk.emit_else(line);

        chunk.emit_op_u16(Op::LOCAL_GET, acc, line);
        chunk.emit_op_u16(Op::LOCAL_GET, base + i, line);
        chunk.emit_call(concat, 2, line);
        chunk.emit_op_u16(Op::LOCAL_SET, acc, line);

        chunk.emit_end(line);
    }

    chunk.emit_op_u16(Op::LOCAL_GET, acc, line);
}

/// PowerShell's `+`, which is three operations chosen by the LEFT operand:
///
/// | `$a` | `$a + $b` |
/// |------|-----------|
/// | array | `$b`'s elements appended — `@(1,2) + 3` is `@(1,2,3)` |
/// | string | concatenation — `'5' + 5` is `'55'` |
/// | number | arithmetic — `5 + '5'` is `10` |
///
/// No shared primitive answers this. `F64_ADD` and `dynamic_add` both coerce an
/// array operand to a number (`@(1,2) + 3` became `NaN`, and `$a += $x` — the
/// idiomatic way to grow an array in PowerShell — trapped in
/// `wasm:js-number.toF64`). `[builtin_slots.array] add` is not a way out: the
/// shared `compile_binop` consults it for EVERY `+`, so binding
/// `collections.concat` there made `10 + 5` evaluate to `5`.
///
/// The left-operand rule is why this cannot be `dynamic_add` either: that
/// concatenates whenever EITHER side is a string, so it answers `'55'` for
/// `5 + '5'` where PowerShell answers `10`.
///
/// Stack: `[a, b]` → `[result]`.
pub fn emit_add(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let b_slot = chunk.alloc_scratch(1);
    let a_slot = chunk.alloc_scratch(1);

    let is_array = chunk.add_import("ecma:array", "isArray");
    let cast_bool = chunk.add_import("wasm:js-boolean", "cast");
    let concat = chunk.add_import("ecma:array", "concat");
    let is_string = chunk.add_import("wasm:js-string", "test");

    chunk.emit_op_u16(Op::LOCAL_SET, b_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, a_slot, line);

    // `isArray` answers with a boxed boolean; the branch needs an i32.
    chunk.emit_op_u16(Op::LOCAL_GET, a_slot, line);
    chunk.emit_call(is_array, 1, line);
    chunk.emit_call(cast_bool, 1, line);
    chunk.emit_if_value(line);

    // Array on the left: append. `concat` flattens an array operand and
    // appends a scalar one, which is exactly what `+` means here.
    chunk.emit_op_u16(Op::LOCAL_GET, a_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, b_slot, line);
    chunk.emit_call(concat, 2, line);

    chunk.emit_else(line);

    // String on the left: concatenate, coercing the right. Note this tests only
    // the LEFT operand — `emit_dyn_add` would concatenate whenever EITHER side
    // is a string and so answer `'55'` for `5 + '5'`, where PowerShell answers
    // `10` because the left operand is a number.
    chunk.emit_op_u16(Op::LOCAL_GET, a_slot, line);
    chunk.emit_call(is_string, 1, line);
    chunk.emit_if_value(line);

    // `emit_dyn_add` concatenates whenever either operand is a string, and the
    // left one is — so here it is exactly right, coercions included.
    chunk.emit_op_u16(Op::LOCAL_GET, a_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, b_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);

    chunk.emit_else(line);

    // Number on the left: arithmetic. `F64_ADD` coerces BOTH operands through
    // `Value::as_f64`, which is what makes `5 + '5'` equal `10`.
    chunk.emit_op_u16(Op::LOCAL_GET, a_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, b_slot, line);
    chunk.emit_op(Op::F64_ADD, line);

    chunk.emit_end(line);
    chunk.emit_end(line);
}

/// `$a.CompareTo($b)` — `System.Object`'s comparison, which every PowerShell
/// value has: `(5).CompareTo(3)` is `1`, `('b').CompareTo('a')` is `1`, and
/// equal values give `0`.
///
/// ⛔**NOT `object.compare`.** That primitive's stack is `[a, b, comparator]`
/// and it INVOKES the comparator unconditionally — there is no null-comparator
/// branch, despite the name. Binding a two-operand `[value_methods]` row to it
/// read a garbage third operand (both examples above answered `-1`), and
/// supplying an explicit `ref.null` only moved the failure to
/// `null is not callable`. It is the comparator-SUPPLIED form of the operation,
/// not the default one.
///
/// The default ordering is built from the shared dynamic comparison ops
/// instead, so `<` and `>` mean here exactly what they mean everywhere else in
/// the language — including for strings, where they are ordinal.
///
/// Stack: `[a, b]` → `[i32]`.
pub fn emit_compare_to(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let b = chunk.alloc_scratch(2);
    let a = b + 1;
    chunk.emit_op_u16(Op::LOCAL_SET, b, line);
    chunk.emit_op_u16(Op::LOCAL_SET, a, line);

    chunk.emit_op_u16(Op::LOCAL_GET, a, line);
    chunk.emit_op_u16(Op::LOCAL_GET, b, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    chunk.emit_i32_const(-1, line);
    chunk.emit_else(line);

    chunk.emit_op_u16(Op::LOCAL_GET, a, line);
    chunk.emit_op_u16(Op::LOCAL_GET, b, line);
    vybe_compiler::primitives::ops::emit_dyn_gt(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    chunk.emit_i32_const(1, line);
    chunk.emit_else(line);
    chunk.emit_i32_const(0, line);
    chunk.emit_end(line);

    chunk.emit_end(line);
}

/// `[char]$x` — the cast to `System.Char`, which is TWO conversions chosen by
/// the operand's runtime type:
///
/// | `$x` | `[char]$x` |
/// |------|------------|
/// | number | the code point — `[char]65` is `A` |
/// | string | the character itself — `[char]'H'` is `H` |
///
/// The slot was bound straight to `strings.from_char_code`, which is only the
/// first row. `[char]'H'` therefore asked for the code point OF A STRING, got a
/// non-number, and rendered as a SPACE — so `[char]::IsUpper([char]'H')` was
/// false and every character-classification test failed on the CAST rather than
/// on the classification it was written to check.
///
/// ⛔The two directions are already declared as a matched pair
/// (`[builtin_slots.char] int` / `[builtin_slots.int] char`) precisely so a
/// cast cannot resolve one way and not the other. That pairing held; what was
/// missing is that ONE of the directions has two source types.
///
/// Stack: `[value]` → `[char]`.
pub fn emit_to_char(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let value = chunk.alloc_scratch(1);
    let is_string = chunk.add_import("wasm:js-string", "test");
    let from_code = chunk.add_import("wasm:js-string", "fromCharCode");

    chunk.emit_op_u16(Op::LOCAL_SET, value, line);

    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    chunk.emit_call(is_string, 1, line);
    chunk.emit_if_value(line);

    // Already a character: a PowerShell `[char]` IS a one-character string.
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);

    chunk.emit_else(line);

    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    chunk.emit_call(from_code, 1, line);

    chunk.emit_end(line);
}

/// `[int]$x` — PowerShell's integer cast, which is THREE conversions:
///
/// | `$x` | `[int]$x` |
/// |------|-----------|
/// | number | banker's rounding — `[int]3.5` is **4**, `[int]2.5` is **2** |
/// | `"65"` | 65 — a numeric parse |
/// | `'A'`  | 65 — the CODE POINT |
///
/// PowerShell tells the last two apart by TYPE: `[int][char]'A'` is the code
/// point and `[int]'A'` on a genuine string is an error. We represent a
/// `[char]` as a one-character string, so the type that would have decided is
/// gone by the time the cast runs and the question has to be asked of the
/// VALUE instead.
///
/// ⛔`char_arithmetic_offset` looked like a REGRESSION from the `[char]` fix
/// and was not: `[char]'A'` used to yield a SPACE and `[int]` of it yielded
/// NaN, so `[char]([int]$ch + 1)` and `[char]'B'` were BOTH a space and the
/// test passed by comparing two wrong answers. Fixing the cast made it fail
/// honestly. Gate a cast change on the whole char suite, never on the one test
/// that names it.
///
/// ⛔The midpoint policy is the shared `MidpointPolicy::HalfEven` primitive,
/// not a hand-rolled `floor(x + 0.5)`. The old lowering answered `3` for
/// `[int]3.5` where PowerShell answers `4`.
///
/// Stack: `[value]` → `[number]`.
pub fn emit_to_int(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let value = chunk.alloc_scratch(2);
    let parsed = value + 1;
    let is_string = chunk.add_import("wasm:js-string", "test");
    let to_number = chunk.add_import("ecma:number", "Number");
    let is_nan = chunk.add_import("ecma:number", "isNaN");
    let char_code = chunk.add_import("wasm:js-string", "charCodeAt");

    chunk.emit_op_u16(Op::LOCAL_SET, value, line);

    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    chunk.emit_call(is_string, 1, line);
    chunk.emit_if_value(line);

    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    chunk.emit_call(to_number, 1, line);
    chunk.emit_op_u16(Op::LOCAL_SET, parsed, line);

    chunk.emit_op_u16(Op::LOCAL_GET, parsed, line);
    chunk.emit_call(is_nan, 1, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    // Not numeric — read it as a character.
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_call(char_code, 2, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, parsed, line);
    vybe_compiler::primitives::math::emit_round(
        chunk,
        vybe_ast::MidpointPolicy::HalfEven,
        line,
    );
    chunk.emit_end(line);

    chunk.emit_else(line);

    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    chunk.emit_call(to_number, 1, line);
    vybe_compiler::primitives::math::emit_round(
        chunk,
        vybe_ast::MidpointPolicy::HalfEven,
        line,
    );

    chunk.emit_end(line);
}

/// `$obj.$name` — a computed member read, where the member is the VALUE held by
/// `$name`.
///
/// ⛔`ExprKind::Index` is the obvious node and the wrong one. Indexing an
/// OBJECT with a non-literal key falls through to ARRAY indexing, so `$h['A']`
/// worked (a literal key compiles to a member read) while `$h[$k]` answered
/// EMPTY for the same key. The shared `[builtin_slots.map] get_item` binding
/// exists but never fires here, because a PowerShell `@{…}` walks to
/// `ExprKind::Object` and is not classified as a Map.
///
/// PowerShell member lookup is CASE-INSENSITIVE — `$obj.UNIQUECODE` and
/// `$obj.uniquecode` both find `UniqueCode` — and the compiler already folds
/// declared member names down, so the key is folded here to match. A static
/// `$obj.UNIQUECODE` gets that fold at compile time; a computed one has to ask
/// for it at runtime.
///
/// Stack: `[obj, key]` → `[value_or_null]`.
pub fn emit_member_dyn(chunks: &mut [Chunk], current: usize, line: u32) {
    vybe_compiler::primitives::strings::emit_to_lower(&mut chunks[current], line);
    vybe_compiler::primitives::dict::emit_get_dynamic(chunks, current, line);
}

/// `$h[$k] = $v` / `$obj.$prop = $v` — a computed WRITE, dispatched on the
/// receiver's runtime type.
///
/// ⛔The mirror of `emit_member_dyn`, and it failed the same way: an `Index`
/// assignment resolves its key at COMPILE time, so `$h['A'] = 1` wrote
/// correctly while `$h[$k] = 2` wrote NOWHERE — silently, with no error and no
/// trap. Reading it back gave empty, which reads as "the write was lost"
/// rather than "the write never happened".
///
/// An array keeps `ecma:array:set` so `$a[$i] = v` is untouched; anything else
/// is a property write under the folded key, matching the read path's fold.
///
/// Stack: `[obj, key, value]` → `[]`.
pub fn emit_index_set(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let value = chunk.alloc_scratch(3);
    let key = value + 1;
    let obj = value + 2;
    let is_array = chunk.add_import("ecma:array", "isArray");
    let cast_bool = chunk.add_import("wasm:js-boolean", "cast");
    let array_set = chunk.add_import("ecma:array", "set");

    chunk.emit_op_u16(Op::LOCAL_SET, value, line);
    chunk.emit_op_u16(Op::LOCAL_SET, key, line);
    chunk.emit_op_u16(Op::LOCAL_SET, obj, line);

    chunk.emit_op_u16(Op::LOCAL_GET, obj, line);
    chunk.emit_call(is_array, 1, line);
    chunk.emit_call(cast_bool, 1, line);
    chunk.emit_if(line);

    chunk.emit_op_u16(Op::LOCAL_GET, obj, line);
    chunk.emit_op_u16(Op::LOCAL_GET, key, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    chunk.emit_call(array_set, 3, line);
    chunk.emit_op(Op::DROP, line);

    chunk.emit_else(line);

    chunk.emit_op_u16(Op::LOCAL_GET, obj, line);
    chunk.emit_op_u16(Op::LOCAL_GET, key, line);
    vybe_compiler::primitives::strings::emit_to_lower(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    vybe_compiler::primitives::dict::emit_set_dynamic(chunks, current, line);

    chunks[current].emit_end(line);
}

/// `$h[$k]` — a computed index READ, dispatched on the receiver's runtime type.
/// The read half of `emit_index_set`; see that comment for why a literal key
/// worked and a variable one did not.
///
/// Stack: `[obj, key]` → `[value]`.
pub fn emit_index_get(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let key = chunk.alloc_scratch(2);
    let obj = key + 1;
    let is_array = chunk.add_import("ecma:array", "isArray");
    let cast_bool = chunk.add_import("wasm:js-boolean", "cast");
    let array_get = chunk.add_import("ecma:array", "get");

    chunk.emit_op_u16(Op::LOCAL_SET, key, line);
    chunk.emit_op_u16(Op::LOCAL_SET, obj, line);

    chunk.emit_op_u16(Op::LOCAL_GET, obj, line);
    chunk.emit_call(is_array, 1, line);
    chunk.emit_call(cast_bool, 1, line);
    chunk.emit_if_value(line);

    chunk.emit_op_u16(Op::LOCAL_GET, obj, line);
    chunk.emit_op_u16(Op::LOCAL_GET, key, line);
    chunk.emit_call(array_get, 2, line);

    chunk.emit_else(line);

    chunk.emit_op_u16(Op::LOCAL_GET, obj, line);
    chunk.emit_op_u16(Op::LOCAL_GET, key, line);
    vybe_compiler::primitives::strings::emit_to_lower(chunk, line);
    vybe_compiler::primitives::dict::emit_get_dynamic(chunks, current, line);

    chunks[current].emit_end(line);
}

/// `… | Out-Null` — evaluate the pipeline, discard its value.
///
/// Stack: `[value]` → `[null]`.
pub fn emit_out_null(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    chunk.emit_op(Op::DROP, line);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

/// One argument of an expanded splat: `$h` by NAME when it is a hashtable,
/// `$arr` by INDEX when it is an array.
///
/// ⛔PowerShell decides between the two by the splat's RUNTIME shape, not by
/// the call site — `F @x` is positional if `$x` is an array and by-name if it
/// is a hashtable — so both spellings reach the callee through this one helper.
/// Parameter names are matched case-insensitively, which
/// `splatting_case_insensitive_parameter_names` asserts (`@{ username = … }`
/// binding `$UserName`).
///
/// Stack: `[container, name, index]` → `[value]`.
pub fn emit_splat_arg(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let index = chunk.alloc_scratch(3);
    let name = index + 1;
    let container = name + 1;
    let is_array = chunk.add_import("ecma:array", "isArray");
    let cast_bool = chunk.add_import("wasm:js-boolean", "cast");
    let array_get = chunk.add_import("ecma:array", "get");

    chunk.emit_op_u16(Op::LOCAL_SET, index, line);
    chunk.emit_op_u16(Op::LOCAL_SET, name, line);
    chunk.emit_op_u16(Op::LOCAL_SET, container, line);

    chunk.emit_op_u16(Op::LOCAL_GET, container, line);
    chunk.emit_call(is_array, 1, line);
    chunk.emit_call(cast_bool, 1, line);
    chunk.emit_if_value(line);

    chunk.emit_op_u16(Op::LOCAL_GET, container, line);
    chunk.emit_op_u16(Op::LOCAL_GET, index, line);
    chunk.emit_call(array_get, 2, line);

    chunk.emit_else(line);

    chunk.emit_op_u16(Op::LOCAL_GET, container, line);
    chunk.emit_op_u16(Op::LOCAL_GET, name, line);
    vybe_compiler::primitives::strings::emit_to_lower(chunk, line);
    vybe_compiler::primitives::dict::emit_get_dynamic(chunks, current, line);

    chunks[current].emit_end(line);
}

/// PowerShell's `/`, which THROWS on a zero divisor rather than answering
/// infinity: `1 / 0` raises `RuntimeException` and is the idiomatic way a test
/// provokes an error for `trap` and `try`/`catch` to catch.
///
/// ⛔This is why the whole `exceptions_trap_statement_scope` family failed with
/// the trap never firing: `1 / 0` evaluated to `Infinity`, so nothing was
/// thrown and the function returned `Infinity AfterTrap` — two success-stream
/// values instead of one. The traps themselves were fine.
///
/// The exception is built with the SHARED `errors` primitives
/// (`emit_class_alloc` → `emit_exception_new_finalize` →
/// `emit_stamp_exception_ancestors` → `emit_throw`), so it carries the same
/// language-exception tag every `try_table` matches on and the ancestor stamp
/// that makes `catch [System.DivideByZeroException]` and its base types match.
///
/// ⚠ Known deviation: PowerShell decides by TYPE, so `[double]1 / 0` is
/// `Infinity` there and throws here. Both operands are `f64` by the time this
/// runs, so the distinction is not available — and no corpus test spells it.
///
/// Stack: `[a, b]` → `[quotient]`, or diverges.
pub fn emit_divide(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let b = chunk.alloc_scratch(2);
    let a = b + 1;
    chunk.emit_op_u16(Op::LOCAL_SET, b, line);
    chunk.emit_op_u16(Op::LOCAL_SET, a, line);

    chunk.emit_op_u16(Op::LOCAL_GET, b, line);
    chunk.emit_f64_const(0.0, line);
    chunk.emit_op(Op::F64_EQ, line);
    chunk.emit_if(line);

    let msg = chunk.alloc_scratch(1);
    chunk.emit_string_const("Attempted to divide by zero.", line);
    chunk.emit_op_u16(Op::LOCAL_SET, msg, line);
    // The shape `emit_exception_new_finalize` expects: [obj, obj, msg].
    vybe_compiler::primitives::class_slots::emit_class_alloc(chunk, line);
    chunk.emit_dup(line);
    chunk.emit_op_u16(Op::LOCAL_GET, msg, line);
    vybe_compiler::primitives::errors::emit_exception_new_finalize(
        chunk,
        "DivideByZeroException",
        line,
    );
    vybe_compiler::primitives::errors::emit_stamp_exception_ancestors(
        chunk,
        "DivideByZeroException",
        line,
    );
    vybe_compiler::primitives::errors::emit_throw(chunk, line);

    chunk.emit_end(line);

    // ⛔NOT a bare `F64_DIV`. `/` still has to reach a USER operator method —
    // `[timespan]::FromHours(6) / [timespan]::FromHours(2)` is `3.0` through
    // `TimeSpan`'s own operator, and dividing the objects numerically answered
    // nonsense. `emit_rich_arithmetic` is the same helper the shared
    // `emit_rich_binop` uses, so the non-zero path behaves exactly as it did
    // before this adapter existed.
    vybe_compiler::primitives::expressions::emit_rich_arithmetic(
        chunk,
        a,
        b,
        &vybe_ast::protocol_slot_key(vybe_ast::ProtocolSlot::Div),
        |c: &mut Chunk, l: u32| c.emit_op(Op::F64_DIV, l),
        line,
    );
}

/// `.GetEnumerator()` — which iterates VALUES for an array and KEY/VALUE PAIRS
/// for a hashtable.
///
/// ⛔One `[value_methods]` row cannot say both, and it was saying the
/// hashtable answer for everything: `[int[]]@(1,2,3).GetEnumerator()` handed
/// back `0,1 1,2 2,3` — index/value pairs — so a class exposing
/// `[IEnumerator] GetEnumerator() { return $this.Items.GetEnumerator() }`
/// summed pairs and answered `NaN`. That is the whole
/// `classes_enumerable_and_iterator` family, and the class was never the
/// problem.
///
/// An array already IS its own enumeration here, so it is handed straight
/// back; everything else keeps `ecma:object:entries`.
///
/// Stack: `[receiver]` → `[enumerable]`.
pub fn emit_get_enumerator(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let recv = chunk.alloc_scratch(1);
    let is_array = chunk.add_import("ecma:array", "isArray");
    let cast_bool = chunk.add_import("wasm:js-boolean", "cast");
    let entries = chunk.add_import("ecma:object", "entries");

    chunk.emit_op_u16(Op::LOCAL_SET, recv, line);

    chunk.emit_op_u16(Op::LOCAL_GET, recv, line);
    chunk.emit_call(is_array, 1, line);
    chunk.emit_call(cast_bool, 1, line);
    chunk.emit_if_value(line);

    chunk.emit_op_u16(Op::LOCAL_GET, recv, line);

    chunk.emit_else(line);

    chunk.emit_op_u16(Op::LOCAL_GET, recv, line);
    chunk.emit_call(entries, 1, line);

    chunk.emit_end(line);
}

/// A pipeline's value: a stream of ONE hands back that one thing.
///
/// `@($obj) | ForEach-Object { $_ }` is the object, not a one-element array —
/// `$r.Tag` reaches the property in PowerShell and answered EMPTY here, because
/// the fold's result stayed wrapped. The same rule the success stream already
/// applies to a function's output, applied to a pipeline's.
///
/// ⛔Only for a length of exactly ONE. Zero stays an empty array and two or
/// more stay an array, which is what every `$x.Count` in the corpus expects.
///
/// Stack: `[value]` → `[value]`.
pub fn emit_unwrap_single(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let value = chunk.alloc_scratch(1);
    let is_array = chunk.add_import("ecma:array", "isArray");
    let cast_bool = chunk.add_import("wasm:js-boolean", "cast");
    let length = chunk.add_import("ecma:array", "length");
    let get = chunk.add_import("ecma:array", "get");

    chunk.emit_op_u16(Op::LOCAL_SET, value, line);

    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    chunk.emit_call(is_array, 1, line);
    chunk.emit_call(cast_bool, 1, line);
    chunk.emit_if_value(line);

    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    chunk.emit_call(length, 1, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op(Op::I32_EQ, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_call(get, 2, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    chunk.emit_end(line);

    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    chunk.emit_end(line);
}
