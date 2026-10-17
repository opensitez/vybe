# Column Bulk Eligibility Short Circuit

`src/emitter/byte_memory.rs` now evaluates the two contiguous-span bounds proofs
only after stride and overlap prove that bulk copying could preserve the
forward scalar result. The conditional declares its i32 result explicitly.
It does not change pointer/count normalization, zero-count handling, unsigned
overlap comparisons, the i64 full-span proof, or the scalar load/store order.

Nonunit strides and higher-destination dependent overlaps retain scalar copies.
Skipping `memory.size` and arithmetic for those cases has no memory effects.
Eligible out-of-bounds spans still use the scalar fallback so that writes
before a bounds trap remain observable. No SIMD, new operation, heap policy,
or representation conversion is introduced.

## Verification Gate

`tests/byte_column_bulk.rs` adds source-emitter tests independent of C lowering:
Number-backed arguments, unsigned high bytes, both overlap directions, self
copies, zero/negative strides, zero count with invalid pointers, invalid spans,
wraparound addresses, and exact memory after source/destination traps. Every
kernel is exported without Vybe metadata and checked by wasmparser.

The debug workload reports guest instruction counts and median timings over
five untraced runs of 128 calls for strided, dependent-overlap, bulk, and empty
copies. Its baseline is a preserved pre-short-circuit emitter in
`/private/tmp/libc_byte_column_before.rs`; the comparison harness is
`/private/tmp/libc_byte_column_compare.rs`. Common memory setup is included in
instruction totals and timings; compiler work and VM construction are excluded
from timings. No frame-level saving is inferred.

Verification is pending: shared debug runtime artifacts disappeared between
inspection and the focused rustc invocation during Main's build. That attempt
failed before compilation because its runtime rlib no longer existed. No
competing Cargo build, alternate target, clean, or process termination was used.
No new instruction count, timing improvement, or external validation pass is
claimed until matching shared debug artifacts become available.

The previous SDL and string results are historical evidence, not a validation
of this short-circuit patch. Allocation/realloc and heap tests belong to James;
this patch does not edit or compile heap_memory.rs directly.
