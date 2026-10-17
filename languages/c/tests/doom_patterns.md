# Doom C Correctness Gates

For a single sequential root-debug command covering existing C lib/frontend, pattern,
and source-backed record suites, see
[One-Command Correctness Gate](frontend_optimization.md#one-command-correctness-gate).
It recompiles through the normal Cargo graph and keeps full Doom emission
opt-in. It does not substitute CLI exit-code tests for parser or WASM evidence.

Run the C frontend and lowering unit tests as well as the execution matrix:

```sh
cargo test -p vybe_language_c --lib -- --nocapture
```

Ignored profiling measurements are not correctness passes. These gates cover
the listed semantics, not all of C or all Doom gameplay paths.

Run the reduced pattern matrix in debug mode:

```sh
cargo test -p vybe_language_c --test doom_patterns -- --nocapture
```

Each case builds and executes its native C reference (`CC`, or `cc`), compiles
the same fixture through Vybe, rejects dynamic equality instructions and ECMA
abstract-equality calls, validates metadata-free WASM, and executes the fixture's
assertions in a fresh VM. Cases are separate Rust tests so failures do not hide
the remaining results. Compile and guest-execution times are reported separately.

| Test Prefix | Covered Shapes |
| --- | --- |
| `pointer_` | Short, scalar, and record strides; subtraction; negative indexes; byte aliases; nested doubly-linked-list stores; pointer-table base/index/value single evaluation; local, global, and heap storage |
| `memory_` | Copy, fill, overlap, search, reallocation, forwarded destinations, pointer return values, single evaluation |
| `update_` | Compound assignment, prefix/postfix values, narrow integer wrap, scalar reference cells, HUD aliases, indexed ammo updates |
| `arithmetic_` | Unsigned wrap, integer promotions, mixed floating operations, signed byte decoding with postincrement, fixed-point wide intermediates, texture-coordinate wrap |
| `input_` | Keyboard command bits, signed movement, nested command copying |
| `pickup_` | Contact radius, height, health caps, bonuses, removal, flags, managed and heap record behavior |
| `renderer_` | Bounded wall clip/plane span algorithms, signed sentinels, interior opening pointers, consecutive frames, pixel guards and no overdraw |

Select a slice by prefix:

```sh
cargo test -p vybe_language_c --test doom_patterns memory_ -- --nocapture
```

The dedicated `doom_pickup_contact`, `managed_record_memory`, `sdl_adapter`,
`record_places`, `fixed_projection`, and `pointer_memory` integration targets cover additional
subsystem and storage cases. The pickup fixture isolates the normal-player
algorithm and selected health items; it does not exercise full-map block-list
traversal or every pickup kind. Pointer and byte-layout assertions deliberately
avoid comparing native host addresses or native pointer sizes with WASM32.

Run the source-backed pickup traversal gate separately:

```sh
cargo test -p vybe_language_c --test record_places pickup_blocklist_traversal_preserves_links_during_removal -- --nocapture
```

It compiles Doom's actual map routines and checks adjacent-block contact,
head/middle unlinking during iteration, surviving links, and excluded radius
boundaries with local and heap mobjs. The touch callback observes contact and
unlinks the item; health, audio, and thinker lifetime behavior are not exercised.

Run the native-backed wall and record-address gate separately:

```sh
cargo test -p vybe_language_c --test record_places indexed_record_addresses_preserve_singleton_and_array_aliases -- --nocapture
```

It executes Doom's actual block-map and movement routines against vertical,
horizontal, and diagonal walls, from both sides, using local and heap geometry.
Rejected moves must retain coordinates and valid sector/block links. It also
audits dynamic equality and validates metadata-free WASM. This small map does
not establish correctness for every full-game map or rendering transition.

Run the complete translation-unit emission gate separately:

```sh
cargo test -p vybe_language_c --test doom_module_contracts -- --ignored --nocapture
```

That gate checks every emitted Doom chunk for dynamic equality and validates the
complete metadata-free WASM module. It does not establish gameplay correctness.

Run the renderer instruction-count gates separately:

```sh
cargo test -p vybe_language_c --test linear_backing_span -- --nocapture
```

They compare optimized loops with their retained scalar fallback, assert pixel
results, reject dynamic equality in both versions, and report executed instruction
counts. The span/column comparisons also validate metadata-free WASM. These are
instruction-count checks, not full-game frame-rate measurements.

Full-game checks use deterministic demo input and Vybe debugger checkpoints for
player position, health, ammo, item removal, and captured frames. Function/line
sampling identifies hot paths, but is not exact branch coverage. A new lowering
failure should gain a reduced fixture plus an appropriate gameplay checkpoint;
manual play is supplementary evidence.

## Remaining Source-Backed Coverage

The `record_places` transition gates additionally execute actual door/plane state machines, real actor-height
propagation on an empty single-sector block map, and actual selected pickup
effects through deferred thinker destruction. See the corresponding
`tests/c/doom_startup/doom_{door_plane_transitions,sector_height_change,pickup_effects}.md`
files for exact scope. Compilation and WASM validation are not branch coverage.

The independent `door_plane_transitions_match_native`,
`sector_height_propagation_matches_native`,
`pickup_effects_and_deferred_removal_match_native`,
`zero_tic_actor_state_transitions_match_native`, and
`enemy_player_search_termination_matches_native` tests prevent one failure from
hiding the other subsystem results. Run them using the `matches_native` filter.

These paths still need native-backed execution fixtures; the existing gates do
not establish their correctness:

The `renderer_wall_plane_spans` pattern copies bounded algorithms from Doom and
uses a deterministic pixel sink. It is not linked to the full renderer source;
textured tiers, perspective/light mapping, and masked drawing remain outside
that case's coverage.

`record_places::momentum_wall_sliding_matches_native_position_and_links` adds
actual `P_XYMovement -> P_SlideMove -> P_PathTraverse` for one vertical wall,
signed momentum splitting, next-tic friction, positions/momentum and links.
It does not cover corners, oblique walls, sector transitions, or step/dropoff
boundaries. Run it with:

```sh
cargo test -p vybe_language_c --test record_places momentum_wall_sliding_matches_native_position_and_links -- --nocapture
```

- Momentum-driven `P_XYMovement -> P_SlideMove -> P_PathTraverse`: corners,
  signed momentum splitting, two-sided openings, step/dropoff boundaries,
  positions, momentum, sector heights, and surviving links.
- `R_StoreWallRange/R_RenderSegLoop -> R_MakeSpans`: signed clip sentinels,
  interior opening pointers, masked columns, consecutive-frame clip arrays,
  opening offsets, emitted spans, and guarded pixels.
- Pickup effects outside the selected stimpack/ammo paths: armor, weapons,
  keys, powers, rejection at full health/ammo, multiplayer, and respawn.
- Enemy player-search wrap/termination, zero-tic actor-state chains, and
  mutation of the thinker list by callbacks during consecutive ticks.
- Full stair-chain construction and actor-height propagation across multiple
  adjacent sectors, corpse/drop-item crushing, and damaging crushers.

For a live movement check, start in `examples/chocolate-doom` with GUI access:

```sh
VYBE_DEBUG_FOLLOW_THREAD=1 ../../target/debug/vybex --debug vybe_doom_core.c -- -iwad freedoom1.wad -warp 1
```

Set `bf G_BuildTiccmd`, continue through startup, and dispatch
`key keydown ArrowUp 38 ArrowUp`. Resume until the key has been processed;
dispatching while paused only queues input. At a subsequent command-builder
entry, inspect `gamekeydown[173]`, then step out and inspect the caller's
`cmd.forwardmove`. Normal walking produces 25. Use `bf P_MovePlayer`,
`bf P_Thrust`, and `bf P_XYMovement` to follow that command into simulation.
At `P_Thrust`, forward walking supplies `move = 51200` before fixed-point
projection. At `P_XYMovement`, confirm the `mo` address is the player actor,
then compare its position before and after stepping out. For linear record
pointers, use `memory` with offsets established from the emitted field access;
the debugger's structural `p` cannot index a raw numeric address.

Release with `key keyup ArrowUp 38 ArrowUp` and verify the key state clears.
Wait for each debugger pause before inspecting: interactive resume commands
are asynchronous. A command reaching simulation is not proof of wall blocking,
item pickup, or sustained frame rate; those need their own checkpoints.
