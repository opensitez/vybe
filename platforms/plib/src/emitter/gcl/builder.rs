use std::sync::Arc;
use vybe_compiler::primitives::instructions::core_wasm;

use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

use super::{GclClass, GclMethod, GclMethodTarget};
use vybe_compiler::primitives::functions::create_function_chunk;

#[derive(Debug, Clone, Copy)]
pub struct AccessorBinding<'a> {
    pub property_name: &'a str,
    pub chunk_idx: usize }

#[derive(Debug, Clone, Copy)]
pub struct MethodBinding<'a> {
    pub method_name: &'a str,
    pub chunk_idx: usize }

/// Push a string as a pool constant. GCL wrapper chunks must not use
/// `Chunk::emit_string_const` — it registers a `wasm:string-constants`
/// import on the chunk, and any local import shadows same-valued baked
/// `chunks[0]` indices in the import-table normalizer's local-first remap.
fn push_string_const(chunk: &mut Chunk, s: &str, line: u32) {
    chunk.emit_string_const(s, line);
}

pub fn build_setter_chunk(
    class_name: &str,
    property_name: &str,
    set_import_idx: u16,
    size_sync_import_idx: Option<u16>,
) -> Chunk {
    let chunk_name = format!("{}::__set_{}", class_name, property_name.to_lowercase());
    let mut chunk = create_function_chunk(&chunk_name, 2);
    let line = 0u32;

    if let Some(event_name) = event_property_name(property_name) {
        chunk.emit_op_u16(Op::LOCAL_GET, 0, line);
        let control_key = chunk.add_constant(Value::String(Arc::from("__control_name")));
        chunk.emit_struct_field_op(Op::STRUCT_GET, 0, control_key, line);
        push_string_const(&mut chunk, event_name, line);
        chunk.emit_op_u16(Op::LOCAL_GET, 1, line);
        chunk.emit_call(set_import_idx, 3, line);
        chunk.emit_op(Op::DROP, line);
        chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        chunk.emit_op(Op::RETURN, line);
        chunk.local_count = 2;
        return chunk;
    }

    chunk.emit_op_u16(Op::LOCAL_GET, 0, line);
    let host_name = host_property_name(property_name);
    push_string_const(&mut chunk, host_name, line);
    chunk.emit_op_u16(Op::LOCAL_GET, 1, line);
    chunk.emit_call(set_import_idx, 3, line);
    chunk.emit_op(Op::DROP, line);
    if is_client_size_property(property_name) {
        if let Some(import_idx) = size_sync_import_idx {
            chunk.emit_op_u16(Op::LOCAL_GET, 0, line);
            chunk.emit_call(import_idx, 1, line);
            chunk.emit_op(Op::DROP, line);
        }
    }
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunk.emit_op(Op::RETURN, line);
    chunk.local_count = 2;
    chunk
}

pub fn build_getter_chunk(class_name: &str, property_name: &str, get_import_idx: u16) -> Chunk {
    let chunk_name = format!("{}::__get_{}", class_name, property_name.to_lowercase());
    let mut chunk = create_function_chunk(&chunk_name, 1);
    let line = 0u32;
    chunk.emit_op_u16(Op::LOCAL_GET, 0, line);
    let host_name = host_property_name(property_name);
    push_string_const(&mut chunk, host_name, line);
    chunk.emit_call(get_import_idx, 2, line);
    chunk.emit_op(Op::RETURN, line);
    chunk.local_count = 1;
    chunk
}

pub fn build_method_chunk(class_name: &str, method: &GclMethod, import_idx: u16) -> Chunk {
    let chunk_name = format!("{}::{}", class_name, method.name);
    let mut chunk = create_function_chunk(&chunk_name, method.arity);
    let line = 0u32;
    match method.target {
        GclMethodTarget::Host { .. } => {
            for slot in 0..method.arity as u16 {
                chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
            }
            chunk.emit_call(import_idx, method.arity, line);
            chunk.emit_op(Op::RETURN, line);
        }
    }
    chunk.local_count = method.arity as u16;
    chunk
}

pub fn build_constructor_chunk(
    class: &GclClass,
    setters: &[AccessorBinding],
    getters: &[AccessorBinding],
    methods: &[MethodBinding],
    widget_new_idx: Option<u16>,
    new_controls_collection_idx: u16,
    new_components_collection_idx: u16,
) -> Chunk {
    let mut chunk = create_function_chunk(class.name, class.ctor_arity);
    let line = 0u32;
    let arity = class.ctor_arity as u16;
    let this_slot = arity;
    let widget_slot = arity + 1;

    if let Some(parent_name) = class.parent {
        vybe_compiler::primitives::globals::emit_read(&mut chunk, parent_name, line);
        if arity > 0 && class.ctor_arity == 1 && parent_name != "TObject" {
            chunk.emit_op_u16(Op::LOCAL_GET, 0, line);
            chunk.emit_op_u8(Op::CALL_REF, 1, line);
        } else {
            chunk.emit_op_u8(Op::CALL_REF, 0, line);
        }
        chunk.emit_op_u16(Op::LOCAL_SET, this_slot, line);
    } else {
        chunk.emit_struct_new(0, 0, line);
        chunk.emit_op_u16(Op::LOCAL_SET, this_slot, line);
    }

    chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
    push_string_const(&mut chunk, class.name, line);
    let type_key = chunk.add_constant(Value::String(Arc::from("__type")));
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, type_key, line);

    for binding in setters {
        bind_ref(
            &mut chunk,
            this_slot,
            &format!("__set_{}", binding.property_name.to_lowercase()),
            binding.chunk_idx,
            line,
        );
    }
    for binding in getters {
        bind_ref(
            &mut chunk,
            this_slot,
            &format!("__get_{}", binding.property_name.to_lowercase()),
            binding.chunk_idx,
            line,
        );
    }
    for binding in methods {
        bind_ref(
            &mut chunk,
            this_slot,
            binding.method_name,
            binding.chunk_idx,
            line,
        );
    }

    if class.name == "TControl" || class.name == "TWinControl" {
        chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
        chunk.emit_call(new_controls_collection_idx, 1, line);
        let key = chunk.add_constant(Value::String(Arc::from("controls")));
        chunk.emit_struct_field_op(Op::STRUCT_SET, 0, key, line);
    }

    if class.name == "TForm" {
        chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
        chunk.emit_call(new_components_collection_idx, 1, line);
        let key = chunk.add_constant(Value::String(Arc::from("components")));
        chunk.emit_struct_field_op(Op::STRUCT_SET, 0, key, line);
    }

    if matches!(class.name, "TMainMenu" | "TPopupMenu" | "TMenuItem") {
        chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
        chunk.emit_call(new_controls_collection_idx, 1, line);
        let key = chunk.add_constant(Value::String(Arc::from("items")));
        chunk.emit_struct_field_op(Op::STRUCT_SET, 0, key, line);
    }

    if matches!(
        class.name,
        "TRadioGroup" | "TComboBox" | "TListBox" | "TListView" | "TTreeView"
    ) {
        chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
        chunk.emit_call(new_controls_collection_idx, 1, line);
        let key = chunk.add_constant(Value::String(Arc::from("items")));
        chunk.emit_struct_field_op(Op::STRUCT_SET, 0, key, line);
    }

    if class.name == "TMemo" {
        chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
        chunk.emit_call(new_controls_collection_idx, 1, line);
        let key = chunk.add_constant(Value::String(Arc::from("lines")));
        chunk.emit_struct_field_op(Op::STRUCT_SET, 0, key, line);
    }

    if let Some(import_idx) = widget_new_idx {
        for i in 0..arity {
            chunk.emit_op_u16(Op::LOCAL_GET, i, line);
        }
        chunk.emit_call(import_idx, arity as u8, line);
        chunk.emit_op_u16(Op::LOCAL_SET, widget_slot, line);

        for field in ["name", "__control_name", "__control_type"] {
            let key_idx = chunk.add_constant(Value::String(Arc::from(field)));
            chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
            chunk.emit_op_u16(Op::LOCAL_GET, widget_slot, line);
            chunk.emit_struct_field_op(Op::STRUCT_GET, 0, key_idx, line);
            chunk.emit_struct_field_op(Op::STRUCT_SET, 0, key_idx, line);
        }
    }

    chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
    chunk.emit_op(Op::RETURN, line);
    chunk.local_count = if widget_new_idx.is_some() {
        arity + 2
    } else {
        arity + 1
    };
    chunk
}

pub fn emit_install_class_global(
    script_chunk: &mut Chunk,
    class_name: &str,
    ctor_idx: usize,
    line: u32,
) {
    script_chunk.emit_op_u16(Op::REF_FUNC, ctor_idx as u16, line);
    script_chunk.emit(0, line);
    vybe_compiler::primitives::globals::emit_write(script_chunk, class_name, line);

    let lower = class_name.to_lowercase();
    if lower != class_name {
        script_chunk.emit_op_u16(Op::REF_FUNC, ctor_idx as u16, line);
        script_chunk.emit(0, line);
        vybe_compiler::primitives::globals::emit_write(script_chunk, lower.as_str(), line);
    }
}

pub fn build_application_run_chunk(run_application_idx: u16) -> Chunk {
    let mut chunk = create_function_chunk("Application::Run", 1);
    let line = 0u32;
    chunk.emit_op_u16(Op::LOCAL_GET, 0, line);
    let main_form_key = chunk.add_constant(Value::String(Arc::from("__main_form")));
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, main_form_key, line);
    chunk.emit_call(run_application_idx, 1, line);
    chunk.emit_op(Op::RETURN, line);
    chunk.local_count = 1;
    chunk
}

pub fn build_application_exit_chunk(app_exit_idx: u16) -> Chunk {
    let mut chunk = create_function_chunk("Application::Exit", 1);
    let line = 0u32;
    chunk.emit_call(app_exit_idx, 0, line);
    chunk.emit_op(Op::RETURN, line);
    chunk.local_count = 1;
    chunk
}

pub fn build_application_initialize_chunk() -> Chunk {
    let mut chunk = create_function_chunk("Application::Initialize", 1);
    let line = 0u32;
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunk.emit_op(Op::RETURN, line);
    chunk.local_count = 1;
    chunk
}

pub fn build_application_title_setter_chunk() -> Chunk {
    let mut chunk = create_function_chunk("Application::__set_title", 2);
    let line = 0u32;
    chunk.emit_op_u16(Op::LOCAL_GET, 0, line);
    chunk.emit_op_u16(Op::LOCAL_GET, 1, line);
    let title_key = chunk.add_constant(Value::String(Arc::from("__gcl_title")));
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, title_key, line);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunk.emit_op(Op::RETURN, line);
    chunk.local_count = 2;
    chunk
}

pub fn build_application_title_getter_chunk() -> Chunk {
    let mut chunk = create_function_chunk("Application::__get_title", 1);
    let line = 0u32;
    chunk.emit_op_u16(Op::LOCAL_GET, 0, line);
    let title_key = chunk.add_constant(Value::String(Arc::from("__gcl_title")));
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, title_key, line);
    chunk.emit_op(Op::RETURN, line);
    chunk.local_count = 1;
    chunk
}

pub fn build_show_message_chunk(msg_box_idx: u16) -> Chunk {
    let mut chunk = create_function_chunk("ShowMessage", 1);
    let line = 0u32;
    chunk.emit_op_u16(Op::LOCAL_GET, 0, line);
    push_string_const(&mut chunk, "Message", line);
    chunk.emit_call(msg_box_idx, 2, line);
    chunk.emit_op(Op::RETURN, line);
    chunk.local_count = 1;
    chunk
}

pub fn build_message_dlg_chunk(msg_box_idx: u16) -> Chunk {
    let mut chunk = create_function_chunk("MessageDlg", 4);
    let line = 0u32;
    chunk.emit_op_u16(Op::LOCAL_GET, 0, line);
    push_string_const(&mut chunk, "Message", line);
    chunk.emit_call(msg_box_idx, 2, line);
    chunk.emit_op(Op::RETURN, line);
    chunk.local_count = 4;
    chunk
}

pub fn emit_install_function_global(
    script_chunk: &mut Chunk,
    name: &str,
    chunk_idx: usize,
    line: u32,
) {
    script_chunk.emit_op_u16(Op::REF_FUNC, chunk_idx as u16, line);
    script_chunk.emit(0, line);
    vybe_compiler::primitives::globals::emit_write(script_chunk, name, line);

    let lower = name.to_lowercase();
    if lower != name {
        script_chunk.emit_op_u16(Op::REF_FUNC, chunk_idx as u16, line);
        script_chunk.emit(0, line);
        vybe_compiler::primitives::globals::emit_write(script_chunk, lower.as_str(), line);
    }
}

pub fn emit_install_application_global(
    script_chunk: &mut Chunk,
    run_chunk_idx: usize,
    exit_chunk_idx: usize,
    initialize_chunk_idx: usize,
    title_setter_idx: usize,
    title_getter_idx: usize,
    line: u32,
) {
    script_chunk.emit_struct_new(0, 0, line);

    core_wasm::dup(script_chunk, line);
    script_chunk.emit_op_u16(Op::REF_FUNC, run_chunk_idx as u16, line);
    script_chunk.emit(0, line);
    let run_key = script_chunk.add_constant(Value::String(Arc::from("run")));
    script_chunk.emit_struct_field_op(Op::STRUCT_SET, 0, run_key, line);

    core_wasm::dup(script_chunk, line);
    script_chunk.emit_op_u16(Op::REF_FUNC, exit_chunk_idx as u16, line);
    script_chunk.emit(0, line);
    let exit_key = script_chunk.add_constant(Value::String(Arc::from("exit")));
    script_chunk.emit_struct_field_op(Op::STRUCT_SET, 0, exit_key, line);

    core_wasm::dup(script_chunk, line);
    script_chunk.emit_op_u16(Op::REF_FUNC, initialize_chunk_idx as u16, line);
    script_chunk.emit(0, line);
    let initialize_key = script_chunk.add_constant(Value::String(Arc::from("initialize")));
    script_chunk.emit_struct_field_op(Op::STRUCT_SET, 0, initialize_key, line);

    core_wasm::dup(script_chunk, line);
    script_chunk.emit_op_u16(Op::REF_FUNC, title_setter_idx as u16, line);
    script_chunk.emit(0, line);
    let set_title_key = script_chunk.add_constant(Value::String(Arc::from("__set_title")));
    script_chunk.emit_struct_field_op(Op::STRUCT_SET, 0, set_title_key, line);

    core_wasm::dup(script_chunk, line);
    script_chunk.emit_op_u16(Op::REF_FUNC, title_getter_idx as u16, line);
    script_chunk.emit(0, line);
    let get_title_key = script_chunk.add_constant(Value::String(Arc::from("__get_title")));
    script_chunk.emit_struct_field_op(Op::STRUCT_SET, 0, get_title_key, line);

    core_wasm::dup(script_chunk, line);
    vybe_compiler::primitives::globals::emit_write(script_chunk, "Application", line);

    vybe_compiler::primitives::globals::emit_write(script_chunk, "application", line);
}

fn bind_ref(chunk: &mut Chunk, this_slot: u16, key: &str, chunk_idx: usize, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
    chunk.emit_op_u16(Op::REF_FUNC, chunk_idx as u16, line);
    chunk.emit(0, line);
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
    let receiver_key = chunk.add_constant(Value::String(Arc::from("__vybe_method_receiver")));
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, receiver_key, line);
    let key_const = chunk.add_constant(Value::String(Arc::from(key)));
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, key_const, line);
}

/// A VCL property spelling → the CANONICAL control property.
///
/// This is what makes a Pascal form controllable from C#: `TLabel.Caption`
/// and `Label.Text` are the same property of the same control, so they must
/// resolve to the same key. The control stores it once, in the widget; only
/// the spelling differs per language.
pub fn host_property_name(property_name: &str) -> &str {
    match property_name.to_ascii_lowercase().as_str() {
        "caption" => "Text",
        "clientwidth" => "Width",
        "clientheight" => "Height",
        other if other.starts_with("on") => match other {
            "onclick" => "Click",
            "onchange" => "Change",
            "oncreate" => "Create",
            "onclose" => "Close",
            "ontimer" => "Timer",
            _ => property_name },
        _ => property_name }
}

pub fn is_event_property(property_name: &str) -> bool {
    event_property_name(property_name).is_some()
}

fn is_client_size_property(property_name: &str) -> bool {
    matches!(
        property_name.to_ascii_lowercase().as_str(),
        "clientwidth" | "clientheight"
    )
}

fn event_property_name(property_name: &str) -> Option<&'static str> {
    match property_name.to_ascii_lowercase().as_str() {
        "onclick" => Some("click"),
        "onchange" => Some("change"),
        "oncreate" => Some("create"),
        "onclose" => Some("close"),
        "ontimer" => Some("timer"),
        "onkeypress" => Some("keyPress"),
        "onkeydown" => Some("keyDown"),
        "onkeyup" => Some("keyUp"),
        _ => None }
}
