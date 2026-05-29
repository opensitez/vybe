//! Emitter unit tests (migrated from the old `vybe_compiler_common` crate).
//! Each module tests one helper family. VM-execution integration tests that
//! duplicated language-level coverage (`test_stdlib_fallback`) were removed
//! — the per-language suites (`vybex/tests/vb`, `js`, `python`, ...) run
//! every stdlib chunk end-to-end through real compiled programs.
mod test_classes;
mod test_errors;
mod test_expressions;
mod test_functions;
mod test_loops;
mod test_ops;
mod test_stdlib;
mod test_strings;
