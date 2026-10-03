//! WinForms captions inside native label/fieldset chrome, with checkable input state.

use vybe_compiler::primitives::gui::{DOCUMENT_MODULE, HOST_FN_ACTIVE_DOCUMENT};
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

const DOM_MODULE: &str = "web:dom";

fn document(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import(DOCUMENT_MODULE, HOST_FN_ACTIVE_DOCUMENT);
    chunk.emit_call(idx, 0, line);
}

fn child(chunk: &mut Chunk, control: u16, first: bool, line: u32) {
    document(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, control, line);
    let idx = chunk.add_import(DOM_MODULE, if first { "firstChild" } else { "lastChild" });
    chunk.emit_call(idx, 2, line);
}

/// Stack: `[label] -> [value]`, or `[label, value] -> [result]`.
pub fn emit_property(
    chunk: &mut Chunk,
    member: &str,
    setting: bool,
    caption_in_first_child: bool,
    line: u32,
) {
    let value = setting.then(|| chunk.alloc_scratch(1));
    if let Some(slot) = value {
        chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
    }
    let control = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, control, line);
    let input = member != "text" || caption_in_first_child;
    let target = chunk.alloc_scratch(1);
    child(chunk, control, input, line);
    chunk.emit_op_u16(Op::LOCAL_SET, target, line);

    document(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, target, line);
    match (member, setting) {
        ("text", true) => {
            chunk.emit_op_u16(Op::LOCAL_GET, value.unwrap(), line);
            let idx = chunk.add_import(DOM_MODULE, "setTextContent");
            chunk.emit_call(idx, 3, line);
        }
        ("text", false) => {
            let idx = chunk.add_import(DOM_MODULE, "textContent");
            chunk.emit_call(idx, 2, line);
        }
        ("checked", true) => {
            chunk.emit_op_u16(Op::LOCAL_GET, value.unwrap(), line);
            let idx = chunk.add_import(DOCUMENT_MODULE, "setChecked");
            chunk.emit_call(idx, 3, line);
        }
        ("checked", false) => {
            let idx = chunk.add_import(DOCUMENT_MODULE, "checked");
            chunk.emit_call(idx, 2, line);
        }
        ("enabled", true) => {
            chunk.emit_string_const("disabled", line);
            chunk.emit_op_u16(Op::LOCAL_GET, value.unwrap(), line);
            chunk.emit_op(Op::I32_EQZ, line);
            let idx = chunk.add_import(DOM_MODULE, "toggleAttribute");
            chunk.emit_call(idx, 4, line);
        }
        ("enabled", false) => {
            chunk.emit_string_const("disabled", line);
            let idx = chunk.add_import(DOM_MODULE, "getAttribute");
            chunk.emit_call(idx, 3, line);
            chunk.emit_op(Op::REF_IS_NULL, line);
        }
        _ => unreachable!("unknown checkable property"),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn groupbox_caption_targets_legend_not_last_control() {
        let mut groupbox = Chunk::new("groupbox_text");
        emit_property(&mut groupbox, "text", true, true, 1);
        assert!(groupbox.imports.iter().any(|item| item.name == "firstChild"));
        assert!(!groupbox.imports.iter().any(|item| item.name == "lastChild"));

        let mut checkbox = Chunk::new("checkbox_text");
        emit_property(&mut checkbox, "text", true, false, 1);
        assert!(checkbox.imports.iter().any(|item| item.name == "lastChild"));
    }
}
