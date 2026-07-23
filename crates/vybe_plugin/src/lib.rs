//! The Vybe plugin SDK.
//!
//! Every capability provider — a source language, a target platform, the
//! compiler, the host, an LSP — is a [`Plugin`] whose `init` registers what it
//! offers into a [`Framework`] (see [`framework`]). This crate also holds the
//! shared frontend types plugins build against: `profile` (LanguageProfile +
//! TOML parser), `class_normalize` (the language-agnostic class IR), and the
//! `registry` tables. Depends only on `vybe_ast`/`vybe_bytecode`/`serde`, never
//! the compiler — this is the crate a loadable dylib links against.

pub mod class_normalize;
pub mod framework;
pub mod profile;
pub mod registry;

pub use framework::{Framework, Plugin, init_all, init_all_on_vm};
