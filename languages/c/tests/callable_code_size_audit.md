# Callable Code Size Audit

## Saved Module

Read-only audit of `/private/tmp/nyx-doom-current.wasm` (2026-10-13 18:25):
99,199,466 total bytes, 98,399,714 code-section bytes (99.19%), 2,854 defined
functions. All local-declaration encodings total only 9,854 bytes. The largest
20 bodies total 21,610,395 bytes. This artifact predates later pointer repairs;
its sizes are evidence about this saved cohort, not verification of new source.

The structured operator audit finds 191 imported functions. The writer emits
chunk bodies in order, then synthetic resume/entry bodies
(`platforms/wasm/src/writer/code.rs`, `encode_code_section`). The `vybe.abi`
section records 2,851 chunk names in the same order. Therefore:

| WASM Function | Chunk Ordinal | Source Function | Body Bytes | Operators |
| --- | --- | --- | --- | --- |
| 574 | 383 | SaveDefaultCollection | 3,290,205 | 1,245,007 |
| 1341 | 1150 | AM_drawLineCharacter | 2,457,137 | 922,483 |

Names identify measured bodies only; they must not become optimization checks.

## Exact Fan-Out

In function 574, `call_ref` appears 30,960 times: exactly 3,440 occurrences of
each of types 3317 through 3325. These signatures have one through nine
externref parameters and one externref result. Function 1341 has 23,382 such
calls: exactly 2,598 occurrences of each signature.

Carrier struct type 263 accounts for 30,960 reads of each of fields 0, 1 and 4
in function 574. Its other dominant counts are 134,141 ref.cast-non-null,
112,259 any.convert_extern, and 103,186 struct.get instructions. This is not
just numeric load/store boxing.

The corresponding generic writer path is the CALL_REF carrier arm in
`platforms/wasm/src/writer/code.rs`: when no exact call signature is present,
it emits a complete arity ladder from `type_ctx.callables.calls` at every
site. Each case re-reads carrier arity/environment/target and repeats casts,
argument restoration/padding and call_ref. The explicit-signature arm already
avoids this ladder. `platforms/wasm/src/writer/callable.rs::uses_carrier_call`
selects representation from chunk/signature metadata, not a language name.

Evidence files: `/private/tmp/c-doom-code-size-audit.json` and
`/private/tmp/c-doom-callable-expansion-audit.txt`. The selected-body operator
audit uses wasmparser, not byte-pattern matching.

## Bounded Proposal

First candidate: share the complete generic carrier dispatcher once per
caller argument-count/result-count pair, using module-private standard WASM
functions. Callers retain their existing argument/callee staging and rooting;
the helper owns the same validation, host-apply branch, carrier arity checks,
hidden environment argument, missing-argument padding and invalid-carrier
trap. Do not remove alternative arities merely because a source language
usually uses exact calls. Exact-signature/raw-core call sites stay unchanged.

Separately, compiler-owned immutable direct bindings or exact producer/signature
metadata can select existing exact/direct calls. Source type spelling or a C
profile is not a sufficient representation/ownership proof. Replacement
dispatchers, writable bindings and unknown callable producers must retain
generic behavior.

Required gates before implementation claims: same/mismatched callable arities,
host versus guest, receiver/capture environment, missing/extra arguments,
callee and argument effects once, replacement bindings, invalid/null carriers,
typed core calls, discarded and multi-result/tail calls. Compare baseline and
candidate metadata-free WASM validity, VM results, emitted instruction/body
bytes, full-module noEq and actual game execution. No reader, VM, scratch-width
or trap-relaxation change is proposed. No size reduction or FPS result has yet
been measured for a candidate.

## Nested Record Expansion

The source-backed `doom_glow_kernel.c` comparison strengthens the C-local
target: actual p_lights.c T_Glow has only a direction switch, four nested
lightlevel compound assignments and two comparisons. Main reports native
ASan/VM success; Clang's isolated standard WASM body is 315 bytes/152 operators,
whereas the isolated Vybe body is 1,050,583 bytes/3,708 locals. These are
different implementation/ABI outputs, not equivalent whole-program link sets.

Read-only prepared-AST dump from the 18:39 CLI, limited to T_Glow:
`/private/tmp/c-glow-prepared-1839.ast`. Its serialized node markers contain
1,332 ternaries, 4,335 members, 2,431 calls, 1,453 casts and 203 sequences.
In particular, `field: "sector"` appears 534 times, `field: "__base"` 1,874
times, and `__c_record_pointer_value` 183 times. This is expansion before the
writer's additional carrier-call fan-out. The dump's 99,862,730 textual bytes
are debug formatting, not the emitted WASM size.

Concrete C lowering sites:

- walk_assignment's compound arm lowers a separate lhs_read before handing
  the expanded value to carray_struct_field_write.
- carray_struct_field_read's generic fallback clones its pointer into
  cstruct/direct-object/carray/raw-array branches. The packed carray read
  further repeats base/index expressions across each byte of a field.
- typed_struct_pointer_field_write captures a nonidentifier scalar-field
  pointer and registers its logical struct type, but does not transfer a
  proven linear backing fact to that captured local. Logical pointer type
  alone cannot justify a raw numeric address.

Prioritized proposal: capture a computed record owner once *before* building
the read/backing branches, and use the captured owner for both sides of a
compound field update. Preserve the exact managed cstruct/carray/raw-byte/
direct-object fallback; do not replace it with the numeric route based on
glow_t/sector_t or any other source name. Captured locals may retain existing
linear producer facts only when linear_pointer_addr_and_stride proves the
original value; raw pointer-slot loads that can return managed carriers are
not such proof. T_Glow's prepared parameter has type_hint None, so a heap
caller by itself does not prove every invocation linear.

Gate the candidate with scalar-field compound assignments through managed
and heap nested records, a numeric outer record with a managed inner pointer,
side-effecting owner/index/RHS once, aliases, signed narrowing and volatile
access counts. Add an AST growth bound for increasing nesting depth, then
rerun actual glow native/VM/WASM/noEq and full game/module gates with recorded
body bytes/operators. No production compaction change is applied by this audit.
