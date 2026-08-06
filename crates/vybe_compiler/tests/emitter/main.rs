//! Emitter unit tests (migrated from the old `vybe_compiler_common` crate).
//! Each module tests one helper family. VM-execution integration tests that
//! duplicated language-level coverage (`test_runtime_helpers_fallback`) were removed
//! — the per-language suites (`vybex/tests/vb`, `js`, `python`, ...) run
//! every runtime helper end-to-end through real compiled programs.

// Force-link vybex so every plugin's link-time (`inventory`) registration
// survives into this integration binary. Integration tests link the lib
// built WITHOUT cfg(test), so build.rs's dev-dep force-link lines don't
// apply here; without this anchor the unreferenced vybex rlib is dropped
// and e.g. the dotnet platform's numeric-format helper never registers
// (`build_runtime_helpers` panics on an empty registry).
use vybex as _;
mod test_classes;
mod test_errors;
mod test_expressions;
mod test_functions;
mod test_generators;
mod test_loops;
mod test_ops;
mod test_runtime_helpers;
mod test_strings;
