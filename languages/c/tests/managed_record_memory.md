# Managed Record Reset Gate

`managed_record_memset_reset.c` reproduces a C lowering failure, not pickup
algorithm behavior. The former `__c_byte_fill` dispatch handled numeric and
`carray` storage, then treated a `cstruct`/named record as an ordinary array.
Consequently local/global named fields survived a repeated zero `memset`.
Heap and contiguous numeric record-array storage already reset correctly.

The C walker now recognizes a statically typed record pointer, a constant
zero fill byte (including 256), and a constant count equal to its complete
record size. It mutates scalar fields, nested records, and primitive arrays
in place. Existing aliases retain their objects/arrays. Proven-linear storage
still uses the original numeric byte-fill fast path. Nested-record address
lowering uses the existing backing-aware field-pointer helper. An outer
`void*` cast preserves `memset`'s C return type on every lowering branch.

This is not general managed-record byte-addressability. Partial/nonzero fills,
runtime counts, erased `void*` record provenance, and arrays of nested records
are not certified by this change. Serializing a managed scalar-cell address
into a numeric record pointer field also remains a separate issue: the dirty
pointer round-trip failed before reset. The reset oracle uses an allocated
pointer target to isolate memory clearing across both backings.

The native-backed Rust integration gate checks local/global/heap/record-array
storage over three resets each; retained nested, scalar, and array aliases;
pointer NULL; destination single evaluation; and returned address identity.
It also audits dynamic equality and validates metadata-free Wasm. The pickup
fixture deliberately keeps explicit initialization as an independent
algorithm gate. `doom_patterns.rs` belongs to the coordinating agent and is
not edited by this lane.
