# Scalar Cell Pointer Serialization

This fixture retains real local/global scalar targets. Allocated scalar targets
are separately labeled controls, not replacements for the failing cases.
Native C passes all 26 pointer identity, dereference, and bidirectional alias
write combinations. A failed identity check avoids dereferencing the corrupt
pointer; the final assertion still fails the entire fixture.

Both the older compiled C gate and fresh shared C rlib reproduce six identity
failures: direct/nested local and global scalar addresses in heap records,
plus nested explicit-cast local scalar addresses in generic and known-linear
heap records. Local/global managed records pass. Allocated targets pass.
An explicit `(int *)&local_value` control repairs direct heap-field identity
without changing the target, but not nested heap-field identity.

Fresh prepared AST/disassembly:
`/private/tmp/scalar-cell-record-fresh-ast.txt`.
The beginning of `check` captures `Unary(AddrOf, local_value)` into
`__c_raw_ptr_store_value3`. There is no generated local declaration for this
carrier. Disassembly stores it in an implicit global at offset 252. Its next
read calls `__vybe_autoderef` at offsets 310/324, before assigning
`__c_linear_ptr_slot_value4` at offset 587. That slot capture is now the scalar
payload instead of its reference object. The number branch clears the side
slot and emits `i32.store` at offset 1156.

The record field retains its `int *` type and correct address arithmetic. The
lost AST distinction is that the generated capture holds an ordinary C
pointer value, not a reference binding requiring implicit value reads.
The shared compiler intentionally marks an identifier assigned `AddrOf` as a
reference binding; the C-generated capture therefore needs to preserve its
pointer-carrier role before this classification. Merely declaring a local
would not prevent that mark.

Relevant C lowering functions are `linear_raw_pointer_store_expr`,
`linear_pointer_slot_store_expr`, `c_pointer_storage_value`, and the
non-identifier-owner path in `typed_struct_pointer_field_write`.
The direct explicit-cast control avoids the first `AddrOf` assignment marker.
The nested path strips pointer casts during storage normalization and then
captures `AddrOf` again, explaining its remaining failure.

Approved scoped production change: both generated captures now wrap their RHS
in a `void*` Cast, including the slot capture after storage normalization.
The existing compiler cast contract preserves runtime identity while avoiding
classification of the generated assignment as a scalar reference binding.
Numeric pointers still follow the existing number branch; managed pointers
still use the existing side slot. No shared compiler, runtime, or libc edits.
An AST unit test checks both address captures and single evaluation.

The production captures are frozen. Once coherent dependencies reappeared,
direct rustc built `/private/tmp/libvybe_language_c-scalar_address_gate.rlib`
against `target/debug/deps/libvybe_compiler-ccc847f0da01446a.rlib` (22:02),
normal runtime `acf705e4b2b5e37d`, and libc `cf46d1ed045ea830`.
C rlib SHA256:
`1c60a39d70e8f73191e42ee016a97e1ef42eff0280c457a08ee3f3a03eb1b330`.
The fresh regression passes all 26 native and VM combinations, full noEq
auditing of 38 chunks, and metadata-free Wasm validation (1,371,333 bytes).
Log: `/private/tmp/scalar-cell-record-current-gates.log`.
The AST unit test has not run separately. No Cargo build was started.

The separate follow-up `scalar_cell_record_pointer_memory.c` checks direct and
nested scalar addresses through whole-record memcpy/memmove, self-move, and
zero memset. It checks alias writes after copying, clearing one record without
clearing the other's pointer, and reusing then zeroing the destination again.
Local, global, and heap records all pass native assertions. Its dedicated
Rust test includes the same VM, noEq, and metadata-free validator gates.
It currently fails in the first managed-local record memcpy: return-pointer
identity passes, but the destination remains marker 29 with its original
direct/nested pointers instead of marker 17 and the copied addresses.
This happens before heap side-slot copying or clearing can be certified.
With diagnostic stage prints, noEq passes 97 chunks and metadata-free Wasm
validates (1,264,597 bytes), but VM execution fails.
Log: `/private/tmp/scalar-cell-record-memory-current-stage.log`.
No auxiliary-slot propagation or clearing claim is made; no broader fix was
applied. Allocated controls in the original 26-case fixture remain intact.

Evidence:
- `/private/tmp/scalar-cell-record-native.log`: all 26 native identities pass.
- `/private/tmp/scalar-cell-record-fresh-integration.log`: fresh shared C
  reproduces the six failures after metadata-free Wasm validation.
- `/private/tmp/scalar-cell-record-noeq.log`: full equality audit passes 38
  chunks on the existing compiled C cohort (no dynamic equality imports).

The native-backed Rust regression is intentionally not ignored. Its original
dyneq/validator/VM version compiled and reproduced the failure on the fresh
shared C cohort. A subsequent addition audits abstract equality imports too;
that updated test relink was blocked by a disappearing/incompatible Webcore
dependency artifact while the parent build ran. No Cargo build was started.
