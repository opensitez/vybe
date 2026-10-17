//! Single complete index operands, normalized at the existing scoped index hook.
use super::*;
pub(super) fn finish(ast: &mut Ast<'_>, raw: grameo::Span, source: &str) {
    let mut child = None;
    let mut multiple = false;
    loop {
        let id = ast.arena.pop().expect("subscript frame must balance");
        if matches!(ast.arena.get(id), Value::Frame(Rule::subscript)) {
            if id == ast.subscript_enter_frame {
                ast.depth(false);
            }
            break;
        }
        if child.is_some() {
            multiple = true;
        } else {
            child = Some(id);
        }
    }
    if multiple {
        return;
    }
    let Some(id) = child else {
        return;
    };
    let Value::Ready { metadata, .. } = ast.arena.get(id) else {
        return;
    };
    // The legacy item walker interprets any colon in its text as a slice.
    // Keep that existing behavior (including colons inside strings) captured.
    if source[raw.start..raw.end].contains(':')
        || metadata.raw.start < raw.start
        || metadata.raw.end > raw.end
    {
        return;
    }
    if !source[raw.start..metadata.raw.start].trim().is_empty() {
        return;
    }
    let tail = source[metadata.raw.end..raw.end].trim();
    if !tail.is_empty() && tail != "," {
        return;
    }
    ast.arena.edit(
        id,
        |value| {
            let Value::Ready { metadata, .. } = value else {
                unreachable!()
            };
            let previous = *metadata;
            metadata.anchor = Some(raw);
            metadata.anchor_rule = Rule::subscript;
            metadata.sequence = false;
            metadata.contextual = true;
            Undo::Anchor(previous)
        },
        |value, undo| {
            let (Value::Ready { metadata, .. }, Undo::Anchor(previous)) = (value, undo) else {
                unreachable!()
            };
            *metadata = previous;
        },
    );
    ast.arena.push(id);
}
