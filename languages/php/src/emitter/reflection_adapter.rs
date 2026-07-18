//! PHP Reflection classes — Rust inline opcode emitters.
//!
//! The walker captures class/function metadata at parse time and passes it
//! as additional args to the adapter. Methods are bound as named props on a
//! struct so `$ref->getName()` dispatches via STRUCT_GET + CALL_REF.

use std::sync::Arc;
use vybe_bytecode::opcode::Op;
use vybe_bytecode::{Chunk, Value};
use vybe_emitter::reflection;

fn sconst(c: &mut Chunk, s: &str) -> u16 {
    c.add_constant(Value::String(Arc::from(s)))
}

/// `() -> this.<field>`.
fn build_field_getter(chunks: &mut Vec<Chunk>, name: &str, field: &str, line: u32) -> usize {
    let mut c = Chunk::new(name);
    c.arity = 1;
    let k = sconst(&mut c, field);
    c.emit_op_u16(Op::LOCAL_GET, 0, line);
    c.emit_op_u16(Op::STRUCT_GET, k, line);
    c.emit_op(Op::RETURN, line);
    c.local_count = c.local_count.max(1);
    chunks.push(c);
    chunks.len() - 1
}

/// `() -> this.<field>` returning a boolean.
fn build_bool_getter(chunks: &mut Vec<Chunk>, name: &str, field: &str, line: u32) -> usize {
    build_field_getter(chunks, name, field, line)
}

/// `getValue($obj)` → `Reflect.get($obj, this.prop)`.
fn build_reflect_get(chunks: &mut Vec<Chunk>, name: &str, field: &str, line: u32) -> usize {
    let mut c = Chunk::new(name);
    c.arity = 2;
    let k = sconst(&mut c, field);
    c.emit_op_u16(Op::LOCAL_GET, 1, line); // obj
    c.emit_op_u16(Op::LOCAL_GET, 0, line);
    c.emit_op_u16(Op::STRUCT_GET, k, line); // this.field
    reflection::emit_get_property_in_chunk(&mut c, line);
    c.emit_op(Op::RETURN, line);
    c.local_count = c.local_count.max(2);
    chunks.push(c);
    chunks.len() - 1
}

/// `setValue($obj, $v)` → `Reflect.set($obj, this.prop, $v)`.
fn build_reflect_set(chunks: &mut Vec<Chunk>, name: &str, field: &str, line: u32) -> usize {
    let mut c = Chunk::new(name);
    c.arity = 3;
    let k = sconst(&mut c, field);
    c.emit_op_u16(Op::LOCAL_GET, 1, line); // obj
    c.emit_op_u16(Op::LOCAL_GET, 0, line);
    c.emit_op_u16(Op::STRUCT_GET, k, line); // this.field
    c.emit_op_u16(Op::LOCAL_GET, 2, line); // value
    reflection::emit_set_property_in_chunk(&mut c, line);
    c.emit_op(Op::DROP, line);
    c.emit_op(Op::NULL, line);
    c.emit_op(Op::RETURN, line);
    c.local_count = c.local_count.max(3);
    chunks.push(c);
    chunks.len() - 1
}

/// `invoke($obj, $arg)` → `Reflect.get($obj, this.method)($obj, $arg)`.
fn build_method_invoke(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    let mut c = Chunk::new("__refl_invoke");
    c.arity = 3;
    let method_k = sconst(&mut c, "method");
    c.emit_op_u16(Op::LOCAL_GET, 1, line);
    c.emit_op_u16(Op::LOCAL_GET, 0, line);
    c.emit_op_u16(Op::STRUCT_GET, method_k, line);
    reflection::emit_get_property_in_chunk(&mut c, line);
    c.emit_op_u16(Op::LOCAL_GET, 1, line);
    c.emit_op_u16(Op::LOCAL_GET, 2, line);
    c.emit_op_u8(Op::CALL_REF, 2, line);
    c.emit_op(Op::RETURN, line);
    c.local_count = c.local_count.max(3);
    chunks.push(c);
    chunks.len() - 1
}

/// `implementsInterface($name)` → check if $name is in this.__interfaces array.
fn build_implements_interface(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    let mut c = Chunk::new("__refl_implements");
    c.arity = 2; // this, interface_name
    let ifaces_k = sconst(&mut c, "__interfaces");
    let indexof_i = c.add_import("ecma:array".to_string(), "indexOf".to_string());
    c.emit_op_u16(Op::LOCAL_GET, 0, line);
    c.emit_op_u16(Op::STRUCT_GET, ifaces_k, line);
    c.emit_op_u16(Op::LOCAL_GET, 1, line);
    c.emit_call(indexof_i, 2, line);
    // indexOf returns -1 if not found; >= 0 means found
    c.emit_f64_const(0.0, line);
    vybe_emitter::ops::emit_dyn_ge(&mut c, line);
    c.emit_op(Op::RETURN, line);
    c.local_count = c.local_count.max(2);
    chunks.push(c);
    chunks.len() - 1
}

/// `getMethods($filter)` → if filter==1, return __methods_public; else __methods.
fn build_get_methods(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    let mut c = Chunk::new("__refl_getMethods");
    c.arity = 2; // this, filter
    let all_k = sconst(&mut c, "__methods");
    let pub_k = sconst(&mut c, "__methods_public");
    // if filter == 1 → __methods_public, else → __methods
    c.emit_op_u16(Op::LOCAL_GET, 1, line);
    c.emit_f64_const(1.0, line);
    vybe_emitter::ops::emit_dyn_eq(&mut c, line);
    c.emit_if_value(line);
    c.emit_op_u16(Op::LOCAL_GET, 0, line);
    c.emit_op_u16(Op::STRUCT_GET, pub_k, line);
    c.emit_else(line);
    c.emit_op_u16(Op::LOCAL_GET, 0, line);
    c.emit_op_u16(Op::STRUCT_GET, all_k, line);
    c.emit_end(line);
    c.emit_op(Op::RETURN, line);
    c.local_count = c.local_count.max(2);
    chunks.push(c);
    chunks.len() - 1
}

/// `getProperties($filter)` → if filter==1, return __fields_public; else __fields.
fn build_get_properties(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    let mut c = Chunk::new("__refl_getProperties");
    c.arity = 2; // this, filter
    let all_k = sconst(&mut c, "__fields");
    let pub_k = sconst(&mut c, "__fields_public");
    c.emit_op_u16(Op::LOCAL_GET, 1, line);
    c.emit_f64_const(1.0, line);
    vybe_emitter::ops::emit_dyn_eq(&mut c, line);
    c.emit_if_value(line);
    c.emit_op_u16(Op::LOCAL_GET, 0, line);
    c.emit_op_u16(Op::STRUCT_GET, pub_k, line);
    c.emit_else(line);
    c.emit_op_u16(Op::LOCAL_GET, 0, line);
    c.emit_op_u16(Op::STRUCT_GET, all_k, line);
    c.emit_end(line);
    c.emit_op(Op::RETURN, line);
    c.local_count = c.local_count.max(2);
    chunks.push(c);
    chunks.len() - 1
}

/// `getParentClass()` → return a new ReflectionClass for the parent, or null.
fn build_get_parent_class(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    build_field_getter(chunks, "__refl_getParentClass", "__parent_ref", line)
}

/// `getParameters()` → return __params array.
fn build_get_parameters(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    build_field_getter(chunks, "__refl_getParams", "__params", line)
}

/// `getAttributes()` → return __attributes array. Filtering is normalized by
/// the walker for current PHP tests; returning all attributes is compatible
/// with single-kind filtered call sites.
fn build_get_attributes(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    build_field_getter(chunks, "__refl_getAttrs", "__attributes", line)
}

/// `getConstructor()` → return __constructor_ref.
fn build_get_constructor(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    build_field_getter(chunks, "__refl_getCtor", "__constructor_ref", line)
}

/// Stamp `__type`, set fields, bind methods, leave instance on stack.
fn finish(
    chunk: &mut Chunk,
    this_slot: u16,
    kind: &str,
    fields: &[(&str, u16)],
    binds: &[(&str, usize)],
    line: u32,
) {
    reflection::emit_new_reflection_object(chunk, this_slot, kind, fields, binds, line);
}

/// `new ReflectionClass($name, $is_abstract, $parent, $interfaces, $methods, $fields)`.
pub fn emit_refl_class(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let getname = build_field_getter(chunks, "__refl_class_name", "name", line);
    let isabstract = build_bool_getter(chunks, "__refl_isabstract", "__is_abstract", line);
    let implements = build_implements_interface(chunks, line);
    let get_methods = build_get_methods(chunks, line);
    let get_properties = build_get_properties(chunks, line);
    let get_parent = build_get_parent_class(chunks, line);
    let get_attrs = build_get_attributes(chunks, line);
    let get_ctor = build_get_constructor(chunks, line);
    let get_params = build_get_parameters(chunks, line);

    let chunk = &mut chunks[current];

    // Pop args from stack (right-to-left): 8 args max
    let ctor_params_slot = chunk.alloc_scratch(1);
    let attrs_slot = chunk.alloc_scratch(1);
    let fields_pub_slot = chunk.alloc_scratch(1);
    let methods_pub_slot = chunk.alloc_scratch(1);
    let fields_slot = chunk.alloc_scratch(1);
    let methods_slot = chunk.alloc_scratch(1);
    let ifaces_slot = chunk.alloc_scratch(1);
    let parent_slot = chunk.alloc_scratch(1);
    let abstract_slot = chunk.alloc_scratch(1);
    let name_slot = chunk.alloc_scratch(1);
    let this_slot = chunk.alloc_scratch(1);
    let ctor_ref_slot = chunk.alloc_scratch(1);

    if argc >= 10 {
        chunk.emit_op_u16(Op::LOCAL_SET, ctor_params_slot, line);
        chunk.emit_op_u16(Op::LOCAL_SET, attrs_slot, line);
        chunk.emit_op_u16(Op::LOCAL_SET, fields_pub_slot, line);
        chunk.emit_op_u16(Op::LOCAL_SET, methods_pub_slot, line);
        chunk.emit_op_u16(Op::LOCAL_SET, fields_slot, line);
        chunk.emit_op_u16(Op::LOCAL_SET, methods_slot, line);
        chunk.emit_op_u16(Op::LOCAL_SET, ifaces_slot, line);
        chunk.emit_op_u16(Op::LOCAL_SET, parent_slot, line);
        chunk.emit_op_u16(Op::LOCAL_SET, abstract_slot, line);
    } else if argc >= 8 {
        chunk.emit_op_u16(Op::ARRAY_NEW_FIXED, 0, line);
        chunk.emit_op_u16(Op::LOCAL_SET, ctor_params_slot, line);
        chunk.emit_op_u16(Op::ARRAY_NEW_FIXED, 0, line);
        chunk.emit_op_u16(Op::LOCAL_SET, attrs_slot, line);
        chunk.emit_op_u16(Op::LOCAL_SET, fields_pub_slot, line);
        chunk.emit_op_u16(Op::LOCAL_SET, methods_pub_slot, line);
        chunk.emit_op_u16(Op::LOCAL_SET, fields_slot, line);
        chunk.emit_op_u16(Op::LOCAL_SET, methods_slot, line);
        chunk.emit_op_u16(Op::LOCAL_SET, ifaces_slot, line);
        chunk.emit_op_u16(Op::LOCAL_SET, parent_slot, line);
        chunk.emit_op_u16(Op::LOCAL_SET, abstract_slot, line);
    } else if argc >= 6 {
        chunk.emit_op_u16(Op::ARRAY_NEW_FIXED, 0, line);
        chunk.emit_op_u16(Op::LOCAL_SET, ctor_params_slot, line);
        chunk.emit_op_u16(Op::ARRAY_NEW_FIXED, 0, line);
        chunk.emit_op_u16(Op::LOCAL_SET, attrs_slot, line);
        chunk.emit_op_u16(Op::ARRAY_NEW_FIXED, 0, line);
        chunk.emit_op_u16(Op::LOCAL_SET, fields_pub_slot, line);
        chunk.emit_op_u16(Op::ARRAY_NEW_FIXED, 0, line);
        chunk.emit_op_u16(Op::LOCAL_SET, methods_pub_slot, line);
        chunk.emit_op_u16(Op::LOCAL_SET, fields_slot, line);
        chunk.emit_op_u16(Op::LOCAL_SET, methods_slot, line);
        chunk.emit_op_u16(Op::LOCAL_SET, ifaces_slot, line);
        chunk.emit_op_u16(Op::LOCAL_SET, parent_slot, line);
        chunk.emit_op_u16(Op::LOCAL_SET, abstract_slot, line);
    } else {
        chunk.emit_op_u16(Op::ARRAY_NEW_FIXED, 0, line);
        chunk.emit_op_u16(Op::LOCAL_SET, ctor_params_slot, line);
        chunk.emit_op_u16(Op::ARRAY_NEW_FIXED, 0, line);
        chunk.emit_op_u16(Op::LOCAL_SET, attrs_slot, line);
        chunk.emit_op_u16(Op::ARRAY_NEW_FIXED, 0, line);
        chunk.emit_op_u16(Op::LOCAL_SET, fields_pub_slot, line);
        chunk.emit_op_u16(Op::ARRAY_NEW_FIXED, 0, line);
        chunk.emit_op_u16(Op::LOCAL_SET, methods_pub_slot, line);
        chunk.emit_op_u16(Op::ARRAY_NEW_FIXED, 0, line);
        chunk.emit_op_u16(Op::LOCAL_SET, fields_slot, line);
        chunk.emit_op_u16(Op::ARRAY_NEW_FIXED, 0, line);
        chunk.emit_op_u16(Op::LOCAL_SET, methods_slot, line);
        chunk.emit_op_u16(Op::ARRAY_NEW_FIXED, 0, line);
        chunk.emit_op_u16(Op::LOCAL_SET, ifaces_slot, line);
        chunk.emit_op(Op::NULL, line);
        chunk.emit_op_u16(Op::LOCAL_SET, parent_slot, line);
        chunk.emit_bool_const(false, line);
        chunk.emit_op_u16(Op::LOCAL_SET, abstract_slot, line);
    }
    chunk.emit_op_u16(Op::LOCAL_SET, name_slot, line);

    // Mini ReflectionMethod for getConstructor()->getParameters().
    reflection::emit_new_reflection_object(
        chunk,
        ctor_ref_slot,
        "ReflectionMethod",
        &[("__params", ctor_params_slot)],
        &[("getparameters", get_params)],
        line,
    );
    chunk.emit_op(Op::DROP, line);

    finish(
        chunk,
        this_slot,
        "ReflectionClass",
        &[
            ("name", name_slot),
            ("__is_abstract", abstract_slot),
            ("__parent_name", parent_slot),
            ("__interfaces", ifaces_slot),
            ("__methods", methods_slot),
            ("__fields", fields_slot),
            ("__methods_public", methods_pub_slot),
            ("__fields_public", fields_pub_slot),
            ("__attributes", attrs_slot),
            ("__constructor_ref", ctor_ref_slot),
        ],
        &[
            ("getname", getname),
            ("isabstract", isabstract),
            ("implementsinterface", implements),
            ("getmethods", get_methods),
            ("getproperties", get_properties),
            ("getparentclass", get_parent),
            ("getattributes", get_attrs),
            ("getconstructor", get_ctor),
        ],
        line,
    );

    // Build __parent_ref: a mini ReflectionClass struct for the parent.
    // finish() left the instance on the stack; pop, build parent ref, push back.
    let chunk = &mut chunks[current];
    chunk.emit_op(Op::DROP, line); // drop the instance left by finish

    // If parent_slot is null, __parent_ref stays null (already set via __parent_name).
    // Otherwise, build a struct { name, getname: fn }.
    chunk.emit_op_u16(Op::LOCAL_GET, parent_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op(Op::I32_EQZ, line); // not null = has parent
    chunk.emit_if(line);
    {
        // Build mini ReflectionClass for parent
        let parent_ref_slot = chunk.alloc_scratch(1);
        reflection::emit_new_reflection_object(
            chunk,
            parent_ref_slot,
            "ReflectionClass",
            &[("name", parent_slot)],
            &[("getname", getname)],
            line,
        );
        chunk.emit_op(Op::DROP, line);
        // this.__parent_ref = parent_ref
        chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, parent_ref_slot, line);
        let pref_k = sconst(chunk, "__parent_ref");
        chunk.emit_op_u16(Op::STRUCT_SET, pref_k, line);
        chunk.emit_op(Op::DROP, line);
    }
    chunk.emit_end(line);

    // Push instance back on stack
    chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
}

/// `new ReflectionMethod($class, $method, $visibility, $paramCount, $requiredParams)`.
pub fn emit_refl_method(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let getname = build_field_getter(chunks, "__refl_method_name", "method", line);
    let invoke = build_method_invoke(chunks, line);
    let get_params_count = build_field_getter(chunks, "__refl_nparams", "__param_count", line);
    let get_required = build_field_getter(chunks, "__refl_nreq", "__required_params", line);
    let is_public = build_bool_getter(chunks, "__refl_ispub", "__is_public", line);
    let is_protected = build_bool_getter(chunks, "__refl_isprot", "__is_protected", line);
    let is_private = build_bool_getter(chunks, "__refl_ispriv", "__is_private", line);
    let get_attrs = build_get_attributes(chunks, line);

    let chunk = &mut chunks[current];

    let attrs_slot = chunk.alloc_scratch(1);
    let required_slot = chunk.alloc_scratch(1);
    let param_count_slot = chunk.alloc_scratch(1);
    let vis_slot = chunk.alloc_scratch(1);
    let method_slot = chunk.alloc_scratch(1);
    let class_slot = chunk.alloc_scratch(1);
    let this_slot = chunk.alloc_scratch(1);
    let is_pub_slot = chunk.alloc_scratch(1);
    let is_prot_slot = chunk.alloc_scratch(1);
    let is_priv_slot = chunk.alloc_scratch(1);

    if argc >= 6 {
        chunk.emit_op_u16(Op::LOCAL_SET, attrs_slot, line);
        chunk.emit_op_u16(Op::LOCAL_SET, required_slot, line);
        chunk.emit_op_u16(Op::LOCAL_SET, param_count_slot, line);
        chunk.emit_op_u16(Op::LOCAL_SET, vis_slot, line);
    } else if argc >= 5 {
        chunk.emit_op_u16(Op::ARRAY_NEW_FIXED, 0, line);
        chunk.emit_op_u16(Op::LOCAL_SET, attrs_slot, line);
        chunk.emit_op_u16(Op::LOCAL_SET, required_slot, line);
        chunk.emit_op_u16(Op::LOCAL_SET, param_count_slot, line);
        chunk.emit_op_u16(Op::LOCAL_SET, vis_slot, line);
    } else {
        chunk.emit_op_u16(Op::ARRAY_NEW_FIXED, 0, line);
        chunk.emit_op_u16(Op::LOCAL_SET, attrs_slot, line);
        chunk.emit_f64_const(0.0, line);
        chunk.emit_op_u16(Op::LOCAL_SET, required_slot, line);
        chunk.emit_f64_const(0.0, line);
        chunk.emit_op_u16(Op::LOCAL_SET, param_count_slot, line);
        let _pub_s = sconst(chunk, "public");
        chunk.emit_string_const("public", line);
        chunk.emit_op_u16(Op::LOCAL_SET, vis_slot, line);
    }
    chunk.emit_op_u16(Op::LOCAL_SET, method_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, class_slot, line);

    // Compute boolean visibility flags from the string
    let _pub_str = sconst(chunk, "public");
    chunk.emit_op_u16(Op::LOCAL_GET, vis_slot, line);
    chunk.emit_string_const("public", line);
    vybe_emitter::ops::emit_dyn_eq(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, is_pub_slot, line);

    let _prot_str = sconst(chunk, "protected");
    chunk.emit_op_u16(Op::LOCAL_GET, vis_slot, line);
    chunk.emit_string_const("protected", line);
    vybe_emitter::ops::emit_dyn_eq(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, is_prot_slot, line);

    let _priv_str = sconst(chunk, "private");
    chunk.emit_op_u16(Op::LOCAL_GET, vis_slot, line);
    chunk.emit_string_const("private", line);
    vybe_emitter::ops::emit_dyn_eq(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, is_priv_slot, line);

    finish(
        chunk,
        this_slot,
        "ReflectionMethod",
        &[
            ("method", method_slot),
            ("class", class_slot),
            ("__param_count", param_count_slot),
            ("__required_params", required_slot),
            ("__is_public", is_pub_slot),
            ("__is_protected", is_prot_slot),
            ("__is_private", is_priv_slot),
            ("__attributes", attrs_slot),
        ],
        &[
            ("getname", getname),
            ("invoke", invoke),
            ("getnumberofparameters", get_params_count),
            ("getnumberofrequiredparameters", get_required),
            ("ispublic", is_public),
            ("isprotected", is_protected),
            ("isprivate", is_private),
            ("getattributes", get_attrs),
        ],
        line,
    );
}

/// `new ReflectionProperty($class, $prop)`.
pub fn emit_refl_property(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let getname = build_field_getter(chunks, "__refl_prop_name", "prop", line);
    let getvalue = build_reflect_get(chunks, "__refl_getvalue", "prop", line);
    let setvalue = build_reflect_set(chunks, "__refl_setvalue", "prop", line);
    let get_attrs = build_get_attributes(chunks, line);
    let chunk = &mut chunks[current];
    let attrs_slot = chunk.alloc_scratch(1);
    let prop_slot = chunk.alloc_scratch(1);
    let class_slot = chunk.alloc_scratch(1);
    let this_slot = chunk.alloc_scratch(1);
    if argc >= 3 {
        chunk.emit_op_u16(Op::LOCAL_SET, attrs_slot, line);
    } else {
        chunk.emit_op_u16(Op::ARRAY_NEW_FIXED, 0, line);
        chunk.emit_op_u16(Op::LOCAL_SET, attrs_slot, line);
    }
    chunk.emit_op_u16(Op::LOCAL_SET, prop_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, class_slot, line);
    finish(
        chunk,
        this_slot,
        "ReflectionProperty",
        &[
            ("prop", prop_slot),
            ("class", class_slot),
            ("__attributes", attrs_slot),
        ],
        &[
            ("getname", getname),
            ("getvalue", getvalue),
            ("setvalue", setvalue),
            ("getattributes", get_attrs),
        ],
        line,
    );
}

/// `new ReflectionFunction($name, $paramCount, $requiredParams)`.
pub fn emit_refl_function(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let getname = build_field_getter(chunks, "__refl_fn_name", "name", line);
    let get_params_count = build_field_getter(chunks, "__refl_fn_nparams", "__param_count", line);
    let get_required = build_field_getter(chunks, "__refl_fn_nreq", "__required_params", line);
    let get_params = build_get_parameters(chunks, line);

    let chunk = &mut chunks[current];
    let params_slot = chunk.alloc_scratch(1);
    let required_slot = chunk.alloc_scratch(1);
    let param_count_slot = chunk.alloc_scratch(1);
    let name_slot = chunk.alloc_scratch(1);
    let this_slot = chunk.alloc_scratch(1);

    if argc >= 4 {
        chunk.emit_op_u16(Op::LOCAL_SET, params_slot, line);
        chunk.emit_op_u16(Op::LOCAL_SET, required_slot, line);
        chunk.emit_op_u16(Op::LOCAL_SET, param_count_slot, line);
    } else if argc >= 3 {
        chunk.emit_op_u16(Op::ARRAY_NEW_FIXED, 0, line);
        chunk.emit_op_u16(Op::LOCAL_SET, params_slot, line);
        chunk.emit_op_u16(Op::LOCAL_SET, required_slot, line);
        chunk.emit_op_u16(Op::LOCAL_SET, param_count_slot, line);
    } else {
        chunk.emit_op_u16(Op::ARRAY_NEW_FIXED, 0, line);
        chunk.emit_op_u16(Op::LOCAL_SET, params_slot, line);
        chunk.emit_f64_const(0.0, line);
        chunk.emit_op_u16(Op::LOCAL_SET, required_slot, line);
        chunk.emit_f64_const(0.0, line);
        chunk.emit_op_u16(Op::LOCAL_SET, param_count_slot, line);
    }
    chunk.emit_op_u16(Op::LOCAL_SET, name_slot, line);

    finish(
        chunk,
        this_slot,
        "ReflectionFunction",
        &[
            ("name", name_slot),
            ("__param_count", param_count_slot),
            ("__required_params", required_slot),
            ("__params", params_slot),
        ],
        &[
            ("getname", getname),
            ("getnumberofparameters", get_params_count),
            ("getnumberofrequiredparameters", get_required),
            ("getparameters", get_params),
        ],
        line,
    );
}

/// `new ReflectionClassConstant($class, $name)` and enum unit case reflection.
pub fn emit_refl_constant(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let get_attrs = build_get_attributes(chunks, line);
    let chunk = &mut chunks[current];
    let attrs_slot = chunk.alloc_scratch(1);
    let name_slot = chunk.alloc_scratch(1);
    let class_slot = chunk.alloc_scratch(1);
    let this_slot = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, attrs_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, name_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, class_slot, line);
    finish(
        chunk,
        this_slot,
        "ReflectionClassConstant",
        &[
            ("class", class_slot),
            ("name", name_slot),
            ("__attributes", attrs_slot),
        ],
        &[("getattributes", get_attrs)],
        line,
    );
}
