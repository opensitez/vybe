//! C has no language-owned namespace tree.
//!
//! Public C library surfaces live in `platforms/libc` adapter registration.
//! The C profile is only a private lowering table for compiler-generated
//! helpers, slots and intrinsics; scraping it into the global namespace tree
//! would make profile rows a second source of public surface.

pub fn register_namespace_tree() {}
