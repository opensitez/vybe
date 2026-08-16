//! `alloc_scratch` slots are RECLAIMED at the statement boundary.
//!
//! `Chunk::alloc_scratch` is a bump allocator over `local_count`. Scratch is
//! dead the moment the statement that asked for it ends, so a function's frame
//! must be sized to the PEAK a single statement needs — never the sum across
//! statements. Without reclamation a long function body paid for every
//! temporary it ever used: a 2500-assertion `.wast` script lowers into one
//! function and ran the u16 slot space out entirely, panicking inside
//! `alloc_scratch` rather than reusing slots dead for thousands of statements.
//!
//! Only the four huge float files in the official spec suite exercised that,
//! and the spec suite is not part of `tests/wast` — so this is the guard.
//!
//! Asserted as EQUALITY between two body sizes, not as a bound. `< n * k`
//! still passes under a partial leak, which is precisely what a future
//! refactor of the allocator would reintroduce. Equality also means no
//! hardcoded slot count, so an unrelated change to what `slice`/`join` cost
//! moves both numbers together and this test stays quiet.
//!
//! ── SCOPE OF THE GUARANTEE ──────────────────────────────────────────────
//! This covers temporaries taken from `alloc_scratch`. It does NOT cover the
//! ones taken from `define_local`, which raises `scope.next_slot` and so
//! raises the reclamation floor permanently. Measured cost per statement:
//!
//!     r = vals.slice(i, 2).join("-")   0 slots   (reclaimed — asserted here)
//!     r = r + "x"                      0 slots   (reclaimed)
//!     vals[i] = 1                      6 slots   (leaks)
//!     r += vals[i]                     8 slots   (leaks)
//!     r += vals[i].toUpperCase()       8 slots   (leaks)
//!     if (vals[i]) { r += "y"; }       8 slots   (leaks)
//!
//! Indexed access is the family that leaks, at ~8 slots a statement — so a
//! function of ~8000 indexing statements still exhausts the u16 slot space.
//! Reclaiming those means `define_local` distinguishing a compiler temporary
//! from a user binding, which is a separate change and not made here. This
//! test deliberately uses a shape from the covered family: asserting the
//! leaking shapes' current numbers would freeze the defect in place as an
//! expectation.

// Anchor the plugin registrations. An integration test links the lib built
// WITHOUT cfg(test), so without this the vybex rlib is dropped — the same
// reason `tests/emitter/main.rs` carries it. Linking alone is not enough
// though: the registry is populated by `init_registered`, which nothing else
// in this binary calls (`languages/kotlin/tests/jvm_probe.rs` does the same).
use vybex as _;

use vybe_compiler::primitives::Compiler;

fn register_plugins() {
    static ONCE: std::sync::Once = std::sync::Once::new();
    ONCE.call_once(vybe_runtime::init_registered);
}

/// Compile a function whose body is `n` scratch-consuming statements and
/// return the frame size the compiler settled on.
///
/// A method chain is deliberate: `emit_invoke_method` stashes its receiver in
/// `alloc_scratch` slots, and chaining gives more than one live at a time, so
/// the peak this measures is a real peak and not one slot.
fn frame_size_for(n: usize) -> u16 {
    let mut src = String::from("function f(vals) {\n  let r = \"\";\n");
    for i in 0..n {
        src.push_str(&format!("  r = vals.slice({i}, 2).join(\"-\");\n"));
    }
    src.push_str("  return r;\n}\n");

    register_plugins();
    let lang = vybe_compiler::languages::all()
        .into_iter()
        .find(|l| l.name == "js")
        .expect("js language plugin registered");
    let profile =
        vybe_compiler::profile::parse_profile((lang.profile_source)()).expect("js profile parses");
    let module = (lang.parse)(&src).expect("source parses");
    let chunks = Compiler::with_profile(profile)
        .compile(&module)
        .expect("module compiles");
    chunks
        .iter()
        .find(|c| c.name == "f")
        .expect("chunk for `f`")
        .local_count
}

#[test]
fn scratch_locals_are_reused_across_statements() {
    let one = frame_size_for(1);
    let eight = frame_size_for(8);
    let thirty_two = frame_size_for(32);

    // The shape must actually consume scratch, or the equality below holds for
    // the boring reason and this test guards nothing. One statement has to be
    // cheaper than eight: the second statement's temporaries sit above the
    // named locals the first one introduced.
    assert!(
        one < eight,
        "statement shape allocates no scratch (1 stmt = {one} slots, 8 = {eight}); \
         the reclamation assertion below would be vacuous"
    );

    // The actual contract: past the point where the peak is established,
    // adding statements costs NOTHING. Four times the body, same frame.
    assert_eq!(
        eight, thirty_two,
        "scratch leaked across statement boundaries: 8 statements = {eight} slots, \
         32 = {thirty_two}. The frame must be sized to the peak ONE statement \
         needs, not the sum over all of them"
    );
}
