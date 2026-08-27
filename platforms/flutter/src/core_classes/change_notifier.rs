//! `package:flutter/foundation.dart`'s `ChangeNotifier`, as a class.
//!
//! Not a widget and never touches the DOM: a list of callbacks plus a disposed
//! flag. That makes it the plainest demonstration that a Flutter adapter class
//! is an ordinary class — declared as AST, walked through `normalize_class` →
//! `primitives/classes.rs`, dispatched by RECEIVER off its own vtable.

use super::builders::*;
use super::notifier;
use vybe_ast::Statement;

/// `class ChangeNotifier { … }`
pub(super) fn change_notifier() -> Statement {
    let mut members = vec![constructor(vec![], notifier::init_state())];
    members.extend(notifier::members());
    class_decl("ChangeNotifier", members)
}
