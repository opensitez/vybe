//! Per-category Flutter widget adapter modules.
//!
//! Each submodule owns one family of Flutter widgets/types and exposes a
//! `pub(crate) const CLASSES: &[FlutterClass]` slice with its adapter entries
//! (and the `FlutterField` field specs those entries reference). The parent
//! [`catalog`](super::catalog) module concatenates every slice listed in
//! [`ALL_CATEGORIES`] into the single lookup table used by the resolver.

use super::catalog::FlutterClass;

pub mod abstracts;
pub mod animation;
pub mod builders;
pub mod cupertino;
pub mod focus;
pub mod gestures;
pub mod images;
pub mod inputs;
pub mod keys;
pub mod layout;
pub mod material;
pub mod painting;
pub mod scrolling;
pub mod value_types;

/// Every category's `CLASSES` slice, in registration order. `catalog`
/// concatenates these once into the aggregate catalog.
pub(crate) const ALL_CATEGORIES: &[&[FlutterClass]] = &[
    abstracts::CLASSES,
    layout::CLASSES,
    painting::CLASSES,
    scrolling::CLASSES,
    material::CLASSES,
    inputs::CLASSES,
    cupertino::CLASSES,
    gestures::CLASSES,
    builders::CLASSES,
    animation::CLASSES,
    images::CLASSES,
    value_types::CLASSES,
    keys::CLASSES,
    focus::CLASSES,
];
