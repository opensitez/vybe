use std::sync::Arc;

use vybe_bytecode::opcode::Op;
use vybe_bytecode::{Chunk, Value};

use crate::emitter::functions::create_function_chunk;
use super::{GclClass, GclMethod, GclMethodTarget};

#[derive(Debug, Clone, Copy)]
pub struct AccessorBinding<'a> {
    pub property_name: &'a str,
    pub chunk_idx: usize,
}

#[derive(Debug, Clone, Copy)]
pub struct MethodBinding<'a> {
    pub method_name: &'a str,
    pub chunk_idx: usize,
}

pub fn build_setter_chunk(class_name: &str, property_name: &str, set_import_idx: u16) -> Chunk {
    let chunk_name = format!("{}::__set_{}", class_name, property_name.to_lowercase());
    let mut chunk = create_function_chunk(&chunk_name, 2);
    let line = 0u32;

    if let Some(event_name) = event_property_name(property_name) {
        chunk.emit_op_u16(Op::LOCAL_GET, 0, line);
        let control_key = chunk.add_constant(Value::String(Arc::from("__control_name")));
        chunk.emit_op_u16(Op::STRUCT_GET, control_key, line);
        let event_const = chunk.add_constant(Value::String(Arc::from(event_name)));
        chunk.emit_op_u16(Op::CONST, event_const, line);
        chunk.emit_op_u16(Op::LOCAL_GET, 1, line);
        chunk.emit_op_u16(Op::CALL_IMPORT, set_import_idx, line);
        chunk.emit(3, line);
        chunk.emit_op(Op::DROP, line);
        chunk.emit_op(Op::NULL, line);
        chunk.emit_op(Op::RETURN, line);
        chunk.local_count = 2;
        return chunk;
    }

    chunk.emit_op_u16(Op::LOCAL_GET, 0, line);
    let host_name = host_property_name(property_name);
    let prop_const = chunk.add_constant(Value::String(Arc::from(host_name)));
    chunk.emit_op_u16(Op::CONST, prop_const, line);
    chunk.emit_op_u16(Op::LOCAL_GET, 1, line);
    chunk.emit_op_u16(Op::CALL_IMPORT, set_import_idx, line);
    chunk.emit(3, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op(Op::NULL, line);
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
    let prop_const = chunk.add_constant(Value::String(Arc::from(host_name)));
    chunk.emit_op_u16(Op::CONST, prop_const, line);
    chunk.emit_op_u16(Op::CALL_IMPORT, get_import_idx, line);
    chunk.emit(2, line);
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
            chunk.emit_op_u16(Op::CALL_IMPORT, import_idx, line);
            chunk.emit(method.arity, line);
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
        let parent_const = chunk.add_constant(Value::String(Arc::from(parent_name)));
        chunk.emit_op_u16(Op::GLOBAL_GET, parent_const, line);
        if arity > 0 && class.ctor_arity == 1 && parent_name != "TObject" {
            chunk.emit_op_u16(Op::LOCAL_GET, 0, line);
            chunk.emit_op_u8(Op::CALL, 1, line);
        } else {
            chunk.emit_op_u8(Op::CALL, 0, line);
        }
        chunk.emit_op_u16(Op::LOCAL_SET, this_slot, line);
        chunk.emit_op(Op::DROP, line);
    } else {
        chunk.emit_op_u16(Op::STRUCT_NEW, 0, line);
        chunk.emit_op_u16(Op::LOCAL_SET, this_slot, line);
        chunk.emit_op(Op::DROP, line);
    }

    chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
    let type_const = chunk.add_constant(Value::String(Arc::from(class.name)));
    chunk.emit_op_u16(Op::CONST, type_const, line);
    let type_key = chunk.add_constant(Value::String(Arc::from("__type")));
    chunk.emit_op_u16(Op::STRUCT_SET, type_key, line);
    chunk.emit_op(Op::DROP, line);

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
        chunk.emit_op_u16(Op::CALL_IMPORT, new_controls_collection_idx, line);
        chunk.emit(1, line);
        let key = chunk.add_constant(Value::String(Arc::from("controls")));
        chunk.emit_op_u16(Op::STRUCT_SET, key, line);
        chunk.emit_op(Op::DROP, line);
    }

    if class.name == "TForm" {
        chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
        chunk.emit_op_u16(Op::CALL_IMPORT, new_components_collection_idx, line);
        chunk.emit(1, line);
        let key = chunk.add_constant(Value::String(Arc::from("components")));
        chunk.emit_op_u16(Op::STRUCT_SET, key, line);
        chunk.emit_op(Op::DROP, line);
    }

    if matches!(class.name, "TMainMenu" | "TPopupMenu" | "TMenuItem") {
        chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
        chunk.emit_op_u16(Op::CALL_IMPORT, new_controls_collection_idx, line);
        chunk.emit(1, line);
        let key = chunk.add_constant(Value::String(Arc::from("items")));
        chunk.emit_op_u16(Op::STRUCT_SET, key, line);
        chunk.emit_op(Op::DROP, line);
    }

    if let Some(import_idx) = widget_new_idx {
        for i in 0..arity {
            chunk.emit_op_u16(Op::LOCAL_GET, i, line);
        }
        chunk.emit_op_u16(Op::CALL_IMPORT, import_idx, line);
        chunk.emit(arity as u8, line);
        chunk.emit_op_u16(Op::LOCAL_SET, widget_slot, line);
        chunk.emit_op(Op::DROP, line);

        for field in ["name", "__control_name", "__control_type"] {
            let key_idx = chunk.add_constant(Value::String(Arc::from(field)));
            chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
            chunk.emit_op_u16(Op::LOCAL_GET, widget_slot, line);
            chunk.emit_op_u16(Op::STRUCT_GET, key_idx, line);
            chunk.emit_op_u16(Op::STRUCT_SET, key_idx, line);
            chunk.emit_op(Op::DROP, line);
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
    let name_const = script_chunk.add_constant(Value::String(Arc::from(class_name)));
    script_chunk.emit_op_u16(Op::GLOBAL_SET, name_const, line);
    script_chunk.emit_op(Op::DROP, line);

    let lower = class_name.to_lowercase();
    if lower != class_name {
        script_chunk.emit_op_u16(Op::REF_FUNC, ctor_idx as u16, line);
        script_chunk.emit(0, line);
        let lower_const = script_chunk.add_constant(Value::String(Arc::from(lower.as_str())));
        script_chunk.emit_op_u16(Op::GLOBAL_SET, lower_const, line);
        script_chunk.emit_op(Op::DROP, line);
    }
}

pub fn build_application_run_chunk(run_application_idx: u16) -> Chunk {
    let mut chunk = create_function_chunk("Application::Run", 1);
    let line = 0u32;
    chunk.emit_op_u16(Op::CALL_IMPORT, run_application_idx, line);
    chunk.emit(0, line);
    chunk.emit_op(Op::RETURN, line);
    chunk.local_count = 1;
    chunk
}

pub fn build_application_exit_chunk(app_exit_idx: u16) -> Chunk {
    let mut chunk = create_function_chunk("Application::Exit", 1);
    let line = 0u32;
    chunk.emit_op_u16(Op::CALL_IMPORT, app_exit_idx, line);
    chunk.emit(0, line);
    chunk.emit_op(Op::RETURN, line);
    chunk.local_count = 1;
    chunk
}

pub fn emit_install_application_global(
    script_chunk: &mut Chunk,
    run_chunk_idx: usize,
    exit_chunk_idx: usize,
    line: u32,
) {
    script_chunk.emit_op_u16(Op::STRUCT_NEW, 0, line);

    script_chunk.emit_op(Op::DUP, line);
    script_chunk.emit_op_u16(Op::REF_FUNC, run_chunk_idx as u16, line);
    script_chunk.emit(0, line);
    let run_key = script_chunk.add_constant(Value::String(Arc::from("run")));
    script_chunk.emit_op_u16(Op::STRUCT_SET, run_key, line);
    script_chunk.emit_op(Op::DROP, line);

    script_chunk.emit_op(Op::DUP, line);
    script_chunk.emit_op_u16(Op::REF_FUNC, exit_chunk_idx as u16, line);
    script_chunk.emit(0, line);
    let exit_key = script_chunk.add_constant(Value::String(Arc::from("exit")));
    script_chunk.emit_op_u16(Op::STRUCT_SET, exit_key, line);
    script_chunk.emit_op(Op::DROP, line);

    script_chunk.emit_op(Op::DUP, line);
    let app_name = script_chunk.add_constant(Value::String(Arc::from("Application")));
    script_chunk.emit_op_u16(Op::GLOBAL_SET, app_name, line);
    script_chunk.emit_op(Op::DROP, line);

    let lower_name = script_chunk.add_constant(Value::String(Arc::from("application")));
    script_chunk.emit_op_u16(Op::GLOBAL_SET, lower_name, line);
    script_chunk.emit_op(Op::DROP, line);
}

fn bind_ref(chunk: &mut Chunk, this_slot: u16, key: &str, chunk_idx: usize, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
    chunk.emit_op_u16(Op::REF_FUNC, chunk_idx as u16, line);
    chunk.emit(0, line);
    let key_const = chunk.add_constant(Value::String(Arc::from(key)));
    chunk.emit_op_u16(Op::STRUCT_SET, key_const, line);
    chunk.emit_op(Op::DROP, line);
}

fn host_property_name(property_name: &str) -> &str {
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
            _ => property_name,
        },
        _ => property_name,
    }
}

pub fn is_event_property(property_name: &str) -> bool {
    event_property_name(property_name).is_some()
}

fn event_property_name(property_name: &str) -> Option<&'static str> {
    match property_name.to_ascii_lowercase().as_str() {
        "onclick" => Some("click"),
        "onchange" => Some("change"),
        "oncreate" => Some("create"),
        "onclose" => Some("close"),
        "ontimer" => Some("timer"),
        _ => None,
    }
}
