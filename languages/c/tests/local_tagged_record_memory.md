# Local Tagged Record Memory

The local inline definition's complete source body previously entered the
owner's type spelling. Metadata was registered under `buffer`, but the owner
looked like `buffer { char before; char bytes[7]; char after; }`. Consequently
`bytes` initialized as scalar zero, sizeof(bytes) became 8 instead of 7, and
byte reads lost their unsigned numeric type.

`Walker::type_text` now extracts the tag from parsed aggregate children when
the aggregate has a body, excluding member declarations from the owner type.
Outer specifiers remain present; declarator pointer/array processing is
unchanged. Anonymous definitions retain their existing generated-tag path.

After that correction, initial memset succeeded, but bzero(object.bytes + 1)
still used ordinary array addition. `rewrite_memset_value` now forms the
existing carray view for a known value-record one-byte array field plus/minus
an offset. Existing proven-linear paths run first and are unchanged.

The remaining dynamic equality came from generated memchr result-index
comparisons, not byte filling. Only the zero/sentinel comparisons now carry
a double Cast: all indexes in [-1, UINT32_MAX] are exactly representable.
The original index used for pointer-offset arithmetic is unchanged.

Fresh focused fixture `local_tagged_record_bytes.c` passes native, VM, noEq
(12 chunks), and metadata-free Wasm validation (70,605 bytes). It covers a
local tag shadowing a file-scope tag, qualifiers, named record arrays,
anonymous records, retained byte aliases, nonzero fill, and partial zeroing.
Full `memory_backing_matrix.c` passes VM, noEq (79 chunks), and metadata-free
Wasm validation (510,515 bytes), including search/overlap/reallocation.
Logs: `/private/tmp/tagged-field-focused-final.log` and
`/private/tmp/tagged-field-matrix-final.log`.
The unchanged doom_patterns matrix passes all 17 native/noEq/Wasm/VM tests
in 25.34 seconds: `/private/tmp/carson-tagged-current-17matrix.log`.
The final C cohort also passes all 26 scalar-address VM combinations,
noEq (38 chunks), and metadata-free Wasm validation (1,371,333 bytes):
`/private/tmp/scalar-cell-final-field-cohort.log`.

Fresh frozen C rlib: `/private/tmp/libvybe_language_c-tagged_field_gate.rlib`.
SHA256: `addb608017b77418eed8faf3d0b6dac22f136f3952fcb20aeb9e07d3e849fe27`.
Compiler: `target/debug/deps/libvybe_compiler-ccc847f0da01446a.rlib` (23:12).
No shared compiler, VM, host, WASI, or profile changes. New AST unit tests are
present but have not been run separately. Record pointer-field copy/reset
follow-up remains a separate exposed failure, not covered by these fixes.
