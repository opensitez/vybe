# Linear String Copy Audit

## Patch

`src/emitter/string_memory.rs` retains the numeric-backing proof already made
by `strncpy` and `strdup`, then emits the existing standard linear-memory copy
primitive instead of calling the general backing-aware copy helper again.
Managed sources/destinations retain their existing adapters. Payload copy count,
bytes, destination identity, allocation headers, and allocation policy are
unchanged. The following zero-fill or NUL store remains a separate operation.

The general copy helper already uses `memory.copy` for linear/linear storage.
Managed/linear transfers cannot become `memory.copy` without changing their
representation; this patch does not attempt that. Realloc already copies only
`min(old_requested_size, new_requested_size)` in bulk. No in-place shrink or
allocator policy change is included. No SIMD or dependent-column change is
included.

## Debug Measurements

Same shared debug compiler/runtime artifacts for before/after. The before
recipe preserves the original two emitter bodies. Inputs use the C Number-backed
ABI, byte value 129, and a terminator. `strncpy` capacity is length plus three.
Instruction totals include common setup and one call; the differences remain
38 instructions for strncpy and 33 for strdup at all measured lengths.

Untraced timings below are the final serial pair, median of five VM runs with
16 calls each. Compilation, VM construction, and debugger instrumentation are
outside the timer; execution setup, memory growth, scanning, copying, and
allocation are inside it. Other agents' builds remained active.

| Operation | Length | Instructions Before | After | Median Before | After |
| --- | ---: | ---: | ---: | ---: | ---: |
| strncpy | 0 | 109 | 71 | 167.458 us | 135.958 us |
| strncpy | 8 | 260 | 222 | 227.292 us | 202.541 us |
| strncpy | 4096 | 10190 | 10152 | 3.850250 ms | 4.047166 ms |
| strdup | 0 | 137 | 104 | 169.291 us | 148.125 us |
| strdup | 8 | 207 | 174 | 198.917 us | 184.500 us |
| strdup | 4096 | 5854 | 5821 | 2.344625 ms | 2.466667 ms |

Timing variation is substantial, and the large-string medians do not establish
a timing improvement. No frame-time/FPS improvement or reduction of Doom's
13.8% conversion-pair share is claimed. Vega's frame log guided the audit;
this patch does not change rendering-loop arithmetic or memory-copy counts.

## Verification

- `languages/c/tests/libc_bulk.rs`: seven passed, one explicitly ignored
  frontend gate, using fresh shared debug libraries without another Cargo build.
- The same three string-emitter tests pass against both original and current
  emission: high-byte data, overlapping adapter ranges in both directions,
  return identity, UTF-8 managed fallback, owned duplicates/private size header,
  and scan/copy/padding trap phases with exact writes before failure.
- Overlapping strncpy is undefined in C; the direct emitter test preserves the
  existing adapter's snapshot behavior rather than asserting a C guarantee.
- Metadata-free generated modules pass wasmparser validation. Both before and
  after Number-backed strncpy/strdup kernels also pass external
  `wasm-tools validate --features all`.
- `tests/c/doom_startup/libc_linear_string_copy_alias.c` passes native `-O0`.
  The available CLI fails its computed-destination strncpy assertion before
  libc emission: source slicing followed by destination assignment is emitted
  instead of the pointer adapter. The fixture is retained unchanged and its
  integration gate is explicitly ignored pending the frontend owner's fix.

No runtime, shared compiler, writer, web, parser, profile, or record_places
file was edited. No git, clean, alternate target, or competing Cargo build was
used.
