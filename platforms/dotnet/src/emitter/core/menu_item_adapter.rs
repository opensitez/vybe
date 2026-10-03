//! WinForms menu-item captions and dropdown collections on native HTML details.

use vybe_compiler::primitives::gui::{DOCUMENT_MODULE, HOST_FN_ACTIVE_DOCUMENT};
use vybe_runtime::Chunk;

pub fn emit_text(chunk: &mut Chunk, setting: bool, line: u32) {
    super::captioned_control_adapter::emit_property(chunk, "text", setting, true, line);
}

/// Stack: `[menu_item] -> [submenu]`.
pub fn emit_dropdown(chunk: &mut Chunk, line: u32) {
    let document = chunk.add_import(DOCUMENT_MODULE, HOST_FN_ACTIVE_DOCUMENT);
    chunk.emit_call(document, 0, line);
    let last_child = chunk.add_import("web:dom", "lastChild");
    chunk.emit_call(last_child, 2, line);
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn dropdown_targets_nested_menu() {
        let mut chunk = Chunk::new("menu_item_dropdown");
        emit_dropdown(&mut chunk, 1);
        assert!(chunk.imports.iter().any(|item| item.name == "lastChild"));
    }
}
