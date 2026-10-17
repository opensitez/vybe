//! Component-layer binary emission for ordered, fully represented definitions.

use crate::encoding::{WASM_MAGIC, write_leb128_i64, write_leb128_u32, write_name, write_section};
use vybe_ast::canon::{
    ComponentBinary, ComponentBinaryCanonOption, ComponentBinaryDef, CoreExportKind, ResourceRep,
    ValSpec, ComponentBinaryAsyncKind, ComponentBinaryAsyncOperation,
    ComponentInterfaceDecl, ComponentInterfaceRef,
};
use std::borrow::Cow;

#[cfg(test)]
#[path = "component_alias_tests.rs"]
mod alias_tests;

fn primitive_type_byte(spec: &ValSpec) -> Result<u8, String> {
    let ValSpec::Prim(name) = spec else {
        return Err("component binary value type is not yet encodable".into());
    };
    crate::component_format::primitive_byte(name)
        .ok_or_else(||format!("unknown component primitive value type: {name}"))
}

fn write_valtype(out: &mut Vec<u8>, spec: &ValSpec, type_count: u32) -> Result<(), String> {
    match spec {
        ValSpec::Prim(_) => out.push(primitive_type_byte(spec)?),
        ValSpec::Ref(index) if *index < type_count => write_leb128_i64(out, i64::from(*index)),
        ValSpec::Ref(index) => return Err(format!("component type index {index} is not declared")),
        _ => return Err("composite component value type requires a prior type definition".into()),
    }
    Ok(())
}

fn write_optional_valtype(
    out: &mut Vec<u8>,
    spec: Option<&ValSpec>,
    type_count: u32,
) -> Result<(), String> {
    if let Some(spec) = spec {
        out.push(1);
        write_valtype(out, spec, type_count)?;
    } else {
        out.push(0);
    }
    Ok(())
}

fn write_value_type(out: &mut Vec<u8>, spec: &ValSpec, type_count: u32) -> Result<(), String> {
    match spec {
        ValSpec::Prim(_) => out.push(primitive_type_byte(spec)?),
        ValSpec::Record(fields) => {
            out.push(0x72);
            write_leb128_u32(out, checked_len(fields.len(), "record fields")?);
            for (name, ty) in fields {
                write_component_name(out, name)?;
                write_valtype(out, ty, type_count)?;
            }
        }
        ValSpec::Variant(cases) => {
            out.push(0x71);
            write_leb128_u32(out, checked_len(cases.len(), "variant cases")?);
            for (name, ty) in cases {
                write_component_name(out, name)?;
                write_optional_valtype(out, ty.as_ref(), type_count)?;
                out.push(0);
            }
        }
        ValSpec::List(ty) => {
            out.push(0x70);
            write_valtype(out, ty, type_count)?;
        }
        ValSpec::ListFixed(ty, len) => {
            out.push(0x67);
            write_valtype(out, ty, type_count)?;
            write_leb128_u32(out, *len);
        }
        ValSpec::Tuple(fields) => {
            out.push(0x6f);
            write_leb128_u32(out, checked_len(fields.len(), "tuple fields")?);
            for ty in fields {
                write_valtype(out, ty, type_count)?;
            }
        }
        ValSpec::Flags(names) | ValSpec::Enum(names) => {
            out.push(if matches!(spec, ValSpec::Flags(_)) {
                0x6e
            } else {
                0x6d
            });
            write_leb128_u32(out, checked_len(names.len(), "type labels")?);
            for name in names {
                write_component_name(out, name)?;
            }
        }
        ValSpec::Option(ty) => {
            out.push(0x6b);
            write_valtype(out, ty, type_count)?;
        }
        ValSpec::Result(ok, err) => {
            out.push(0x6a);
            write_optional_valtype(out, ok.as_deref(), type_count)?;
            write_optional_valtype(out, err.as_deref(), type_count)?;
        }
        ValSpec::Own(index) | ValSpec::Borrow(index) => {
            if *index >= type_count {
                return Err(format!("component type index {index} is not declared"));
            }
            out.push(if matches!(spec, ValSpec::Own(_)) {
                0x69
            } else {
                0x68
            });
            write_leb128_u32(out, *index);
        }
        ValSpec::Stream(ty) | ValSpec::Future(ty) => {
            out.push(if matches!(spec, ValSpec::Stream(_)) {
                0x66
            } else {
                0x65
            });
            write_optional_valtype(out, ty.as_deref(), type_count)?;
        }
        ValSpec::Map(key, value) => {
            out.push(0x63);
            write_valtype(out, key, type_count)?;
            write_valtype(out, value, type_count)?;
        }
        ValSpec::Ref(_) => {
            return Err("a component type definition cannot be a bare type reference".into());
        }
    }
    Ok(())
}

fn write_type_definition(
    out: &mut Vec<u8>,
    definition: &ComponentBinaryDef,
    type_count: u32,
) -> Result<(), String> {
    match definition {
        ComponentBinaryDef::InterfaceType { is_component, declarations } => {
            out.push(if *is_component { 0x41 } else { 0x42 });
            write_leb128_u32(out, checked_len(declarations.len(), "interface declarations")?);
            let mut local_types = 0_u32;
            for declaration in declarations {
                match declaration {
                    ComponentInterfaceDecl::Type(ty) => {
                        out.push(1);
                        write_type_definition(out, ty, local_types)?;
                        local_types = local_types.checked_add(1).ok_or("interface type index overflow")?;
                    }
                    ComponentInterfaceDecl::AliasOuterType { count, index } => {
                        out.extend_from_slice(&[2, 3, 2]);
                        write_leb128_u32(out, *count);
                        write_leb128_u32(out, *index);
                        local_types = local_types.checked_add(1).ok_or("interface type index overflow")?;
                    }
                    ComponentInterfaceDecl::AliasInstanceExport { instance, name, kind } => {
                        use vybe_ast::canon::ComponentItemKind as K;
                        if !matches!(kind, K::Type | K::Instance) {
                            return Err("interface export alias must have type or instance sort".into());
                        }
                        out.push(2);
                        write_item_sort(out, *kind);
                        out.push(0);
                        write_leb128_u32(out, *instance);
                        write_component_name(out, name)?;
                        if *kind == K::Type {
                            local_types = local_types.checked_add(1).ok_or("interface type index overflow")?;
                        }
                    }
                    ComponentInterfaceDecl::Import { name, ty }
                    | ComponentInterfaceDecl::Export { name, ty } => {
                        let importing = matches!(declaration, ComponentInterfaceDecl::Import { .. });
                        if importing && !is_component {
                            return Err("instance type cannot declare imports".into());
                        }
                        out.push(if importing { 3 } else { 4 });
                        out.push(0);
                        write_component_name(out, name)?;
                        write_interface_ref(out, ty, local_types)?;
                        if matches!(ty, ComponentInterfaceRef::TypeEq(_) | ComponentInterfaceRef::TypeSubResource) {
                            local_types = local_types.checked_add(1).ok_or("interface type index overflow")?;
                        }
                    }
                }
            }
            Ok(())
        }
        ComponentBinaryDef::ValueType(spec) => write_value_type(out, spec, type_count),
        ComponentBinaryDef::FuncType {
            params,
            result,
            is_async,
        } => {
            out.push(if *is_async { 0x43 } else { 0x40 });
            write_leb128_u32(
                out,
                checked_len(params.len(), "component function parameters")?,
            );
            for (name, ty) in params {
                write_component_name(out, name)?;
                write_valtype(out, ty, type_count)?;
            }
            match result {
                Some(ty) => {
                    out.push(0);
                    write_valtype(out, ty, type_count)?;
                }
                None => out.extend_from_slice(&[1, 0]),
            }
            Ok(())
        }
        ComponentBinaryDef::ResourceType { rep, dtor } => {
            out.push(0x3f);
            out.push(match rep {
                ResourceRep::I32 => 0x7f,
                ResourceRep::I64 => 0x7e,
            });
            if let Some(dtor) = dtor {
                out.push(1);
                write_leb128_u32(out, *dtor);
            } else {
                out.push(0);
            }
            Ok(())
        }
        _ => Err("component definition is not a type".into()),
    }
}

fn write_interface_ref(out: &mut Vec<u8>, ty: &ComponentInterfaceRef, type_count: u32) -> Result<(), String> {
    let index = match ty {
        ComponentInterfaceRef::Module(index) => { out.extend_from_slice(&[0, 0x11]); *index }
        ComponentInterfaceRef::Func(index) => { out.push(1); *index }
        ComponentInterfaceRef::Value(value) => {
            out.push(2);
            write_valtype(out, value, type_count)?;
            return Ok(());
        }
        ComponentInterfaceRef::TypeEq(index) => { out.extend_from_slice(&[3, 0]); *index }
        ComponentInterfaceRef::TypeSubResource => { out.extend_from_slice(&[3, 1]); return Ok(()); }
        ComponentInterfaceRef::Component(index) => { out.push(4); *index }
        ComponentInterfaceRef::Instance(index) => { out.push(5); *index }
    };
    write_leb128_u32(out, index);
    Ok(())
}

fn write_item_ref(out: &mut Vec<u8>, item: &vybe_ast::canon::ComponentItemRef, counts: [u32; 6]) -> Result<(), String> {
    use vybe_ast::canon::ComponentItemKind as K;
    let sort = match item.kind {
        K::Module => 0,
        K::Func => 1,
        K::Value => 2,
        K::Type => 3,
        K::Component => 4,
        K::Instance => 5,
    };
    if item.index >= counts[sort] {
        return Err(format!("component {:?} index {} is not declared", item.kind, item.index));
    }
    write_item_sort(out, item.kind);
    write_leb128_u32(out, item.index);
    Ok(())
}

fn write_item_sort(out: &mut Vec<u8>, kind: vybe_ast::canon::ComponentItemKind) {
    use vybe_ast::canon::ComponentItemKind as K;
    match kind {
        K::Module => out.extend_from_slice(&[0, 0x11]),
        K::Func => out.push(1),
        K::Value => out.push(2),
        K::Type => out.push(3),
        K::Component => out.push(4),
        K::Instance => out.push(5),
    }
}

fn write_canon_options(
    out: &mut Vec<u8>,
    options: &[ComponentBinaryCanonOption],
    core_func_count: u32,
    core_memory_count: u32,
) -> Result<(), String> {
    write_leb128_u32(out, checked_len(options.len(), "canonical options")?);
    for option in options {
        let (opcode, index) = match *option {
            ComponentBinaryCanonOption::Utf8 => (0, None),
            ComponentBinaryCanonOption::Utf16 => (1, None),
            ComponentBinaryCanonOption::CompactUtf16 => (2, None),
            ComponentBinaryCanonOption::Memory(index) => {
                if index >= core_memory_count {
                    return Err(format!("core memory index {index} is not declared"));
                }
                (3, Some(index))
            }
            ComponentBinaryCanonOption::Realloc(index) => (4, Some(index)),
            ComponentBinaryCanonOption::PostReturn(index) => (5, Some(index)),
            ComponentBinaryCanonOption::Async => (6, None),
            ComponentBinaryCanonOption::Callback(index) => (7, Some(index)),
            ComponentBinaryCanonOption::CoreType(index) => (8, Some(index)),
            ComponentBinaryCanonOption::Gc => (9, None),
        };
        if matches!(opcode, 4 | 5 | 7) && index.is_some_and(|i| i >= core_func_count) {
            return Err(format!(
                "canonical option core function index {} is not declared",
                index.unwrap()
            ));
        }
        out.push(opcode);
        if let Some(index) = index {
            write_leb128_u32(out, index);
        }
    }
    Ok(())
}

fn write_canon_definition(
    out: &mut Vec<u8>,
    definition: &ComponentBinaryDef,
    type_count: u32,
    core_func_count: &mut u32,
    component_func_count: &mut u32,
    core_memory_count: u32,
    core_type_count: u32,
    core_table_count: u32,
) -> Result<(), String> {
    match definition {
        ComponentBinaryDef::CanonErrorContextNew { options }
        | ComponentBinaryDef::CanonErrorContextDebugMessage { options } => {
            out.push(if matches!(definition, ComponentBinaryDef::CanonErrorContextNew { .. }) {
                0x1c
            } else { 0x1d });
            write_canon_options(out, options, *core_func_count, core_memory_count)?;
            *core_func_count = core_func_count.checked_add(1)
                .ok_or("core function index overflow")?;
        }
        ComponentBinaryDef::CanonErrorContextDrop => {
            out.push(0x1e);
            *core_func_count = core_func_count.checked_add(1)
                .ok_or("core function index overflow")?;
        }
        ComponentBinaryDef::CanonResourceNew { type_index }
        | ComponentBinaryDef::CanonResourceDrop { type_index }
        | ComponentBinaryDef::CanonResourceRep { type_index } => {
            if *type_index >= type_count {
                return Err(format!("component resource type index {type_index} is not declared"));
            }
            out.push(match definition {
                ComponentBinaryDef::CanonResourceNew { .. } => 0x02,
                ComponentBinaryDef::CanonResourceDrop { .. } => 0x03,
                ComponentBinaryDef::CanonResourceRep { .. } => 0x04,
                _ => unreachable!(),
            });
            write_leb128_u32(out, *type_index);
            *core_func_count = core_func_count.checked_add(1)
                .ok_or("core function index overflow")?;
        }
        ComponentBinaryDef::CanonContext { set, ty, slot } => {
            use vybe_ast::canon::ComponentCoreValType as T;
            let byte = match ty {
                T::I32 => 0x7f, T::I64 => 0x7e,
                _ => return Err("canonical context type must be i32 or i64".into()),
            };
            if *slot >= 2 { return Err("canonical context slot must be less than 2".into()); }
            out.extend_from_slice(&[if *set { 0x0b } else { 0x0a }, byte]);
            write_leb128_u32(out, *slot);
            *core_func_count = core_func_count.checked_add(1).ok_or("core function index overflow")?;
        }
        ComponentBinaryDef::CanonTaskCancel
        | ComponentBinaryDef::CanonSubtaskCancel { .. }
        | ComponentBinaryDef::CanonSubtaskDrop
        | ComponentBinaryDef::CanonBackpressureInc
        | ComponentBinaryDef::CanonBackpressureDec => {
            match definition {
                ComponentBinaryDef::CanonTaskCancel => out.push(0x05),
                ComponentBinaryDef::CanonSubtaskCancel { is_async } => out.extend_from_slice(&[0x06, u8::from(*is_async)]),
                ComponentBinaryDef::CanonSubtaskDrop => out.push(0x0d),
                ComponentBinaryDef::CanonBackpressureInc => out.push(0x24),
                ComponentBinaryDef::CanonBackpressureDec => out.push(0x25),
                _ => unreachable!(),
            }
            *core_func_count = core_func_count.checked_add(1).ok_or("core function index overflow")?;
        }
        ComponentBinaryDef::CanonWaitableSetNew
        | ComponentBinaryDef::CanonWaitableSetDrop
        | ComponentBinaryDef::CanonWaitableJoin
        | ComponentBinaryDef::CanonWaitableSetWait { .. }
        | ComponentBinaryDef::CanonWaitableSetPoll { .. } => {
            match definition {
                ComponentBinaryDef::CanonWaitableSetNew => out.push(0x1f),
                ComponentBinaryDef::CanonWaitableSetDrop => out.push(0x22),
                ComponentBinaryDef::CanonWaitableJoin => out.push(0x23),
                ComponentBinaryDef::CanonWaitableSetWait { memory }
                | ComponentBinaryDef::CanonWaitableSetPoll { memory } => {
                    if *memory >= core_memory_count {
                        return Err(format!("core memory index {memory} is not declared"));
                    }
                    out.push(if matches!(definition, ComponentBinaryDef::CanonWaitableSetWait { .. }) {
                        0x20
                    } else { 0x21 });
                    out.push(0);
                    write_leb128_u32(out, *memory);
                }
                _ => unreachable!(),
            }
            *core_func_count = core_func_count.checked_add(1)
                .ok_or("core function index overflow")?;
        }
        ComponentBinaryDef::CanonThreadResumeLater
        | ComponentBinaryDef::CanonThreadSuspend
        | ComponentBinaryDef::CanonThreadYield
        | ComponentBinaryDef::CanonThreadSuspendThenResume
        | ComponentBinaryDef::CanonThreadYieldThenResume
        | ComponentBinaryDef::CanonThreadSuspendThenPromote
        | ComponentBinaryDef::CanonThreadYieldThenPromote => {
            let opcode = match definition {
                ComponentBinaryDef::CanonThreadResumeLater => 0x28,
                ComponentBinaryDef::CanonThreadSuspend => 0x29,
                ComponentBinaryDef::CanonThreadYield => 0x0c,
                ComponentBinaryDef::CanonThreadSuspendThenResume => 0x2a,
                ComponentBinaryDef::CanonThreadYieldThenResume => 0x2b,
                ComponentBinaryDef::CanonThreadSuspendThenPromote => 0x2c,
                ComponentBinaryDef::CanonThreadYieldThenPromote => 0x2d,
                _ => unreachable!(),
            };
            out.push(opcode);
            if opcode != 0x28 {
                out.push(0);
            }
            *core_func_count = core_func_count.checked_add(1)
                .ok_or("core function index overflow")?;
        }
        ComponentBinaryDef::CanonThreadIndex => {
            out.push(0x26);
            *core_func_count = core_func_count.checked_add(1)
                .ok_or("core function index overflow")?;
        }
        ComponentBinaryDef::CanonThreadNewIndirect { type_index, table_index } => {
            if *type_index >= core_type_count {
                return Err(format!("core type index {type_index} is not declared"));
            }
            if *table_index >= core_table_count {
                return Err(format!("core table index {table_index} is not declared"));
            }
            out.push(0x27);
            write_leb128_u32(out, *type_index);
            write_leb128_u32(out, *table_index);
            *core_func_count = core_func_count.checked_add(1)
                .ok_or("core function index overflow")?;
        }
        ComponentBinaryDef::CanonLift {
            core_func,
            type_index,
            options,
        } => {
            if *core_func >= *core_func_count {
                return Err(format!("core function index {core_func} is not declared"));
            }
            if *type_index >= type_count {
                return Err(format!("component type index {type_index} is not declared"));
            }
            out.extend_from_slice(&[0, 0]);
            write_leb128_u32(out, *core_func);
            write_canon_options(out, options, *core_func_count, core_memory_count)?;
            write_leb128_u32(out, *type_index);
            *component_func_count = component_func_count
                .checked_add(1)
                .ok_or("component function index overflow")?;
        }
        ComponentBinaryDef::CanonLower { func, options } => {
            if *func >= *component_func_count {
                return Err(format!("component function index {func} is not declared"));
            }
            out.extend_from_slice(&[1, 0]);
            write_leb128_u32(out, *func);
            write_canon_options(out, options, *core_func_count, core_memory_count)?;
            *core_func_count = core_func_count
                .checked_add(1)
                .ok_or("core function index overflow")?;
        }
        ComponentBinaryDef::CanonTaskReturn { result, options } => {
            out.push(0x09);
            if let Some(result) = result {
                out.push(0);
                write_valtype(out, result, type_count)?;
            } else {
                out.extend_from_slice(&[1, 0]);
            }
            write_canon_options(out, options, *core_func_count, core_memory_count)?;
            *core_func_count = core_func_count
                .checked_add(1)
                .ok_or("core function index overflow")?;
        }
        ComponentBinaryDef::CanonAsync { kind, operation, type_index, options } => {
            use ComponentBinaryAsyncOperation::*;
            if *type_index >= type_count {
                return Err(format!("component type index {type_index} is not declared"));
            }
            let base = match kind {
                ComponentBinaryAsyncKind::Stream => 0x0e,
                ComponentBinaryAsyncKind::Future => 0x15,
            };
            let offset = match operation {
                New => 0, Read => 1, Write => 2, CancelRead { .. } => 3,
                CancelWrite { .. } => 4, DropReadable => 5, DropWritable => 6,
                Forward => match kind {
                    ComponentBinaryAsyncKind::Stream => 0x2e - base,
                    ComponentBinaryAsyncKind::Future => 0x2f - base,
                },
            };
            out.push(base + offset);
            write_leb128_u32(out, *type_index);
            match operation {
                Read | Write => write_canon_options(out, options, *core_func_count, core_memory_count)?,
                _ => {
                    if !options.is_empty() {
                        return Err("canonical async operation does not accept options".into());
                    }
                    if let CancelRead { is_async } | CancelWrite { is_async } = operation {
                        out.push(u8::from(*is_async));
                    }
                }
            }
            *core_func_count = core_func_count.checked_add(1)
                .ok_or("core function index overflow")?;
        }
        _ => return Err("component definition is not canonical".into()),
    }
    Ok(())
}

fn core_kind_byte(kind: CoreExportKind) -> u8 {
    match kind {
        CoreExportKind::Func => 0,
        CoreExportKind::Table => 1,
        CoreExportKind::Memory => 2,
        CoreExportKind::Global => 3,
        CoreExportKind::Tag => 4,
    }
}

fn checked_len(len: usize, what: &str) -> Result<u32, String> {
    u32::try_from(len).map_err(|_| format!("{what} exceeds the component binary u32 limit"))
}

fn write_core_function_type(
    out: &mut Vec<u8>,
    params: &[vybe_ast::canon::ComponentCoreValType],
    results: &[vybe_ast::canon::ComponentCoreValType],
) -> Result<(), String> {
    use vybe_ast::canon::ComponentCoreValType as T;
    out.push(0x60);
    for values in [params, results] {
        write_leb128_u32(out, checked_len(values.len(), "core function values")?);
        for value in values {
            out.push(match value {
                T::I32 => 0x7f,
                T::I64 => 0x7e,
                T::F32 => 0x7d,
                T::F64 => 0x7c,
                T::V128 => 0x7b,
            });
        }
    }
    Ok(())
}

fn write_component_name(out: &mut Vec<u8>, name: &str) -> Result<(), String> {
    checked_len(name.len(), "component name")?;
    write_name(out, name);
    Ok(())
}

// Large payloads stay borrowed across nested sections until the root is flattened.
enum ComponentOutput<'a> {
    Contiguous(Vec<u8>),
    Segmented { parts: Vec<Cow<'a, [u8]>>, len: usize },
}

impl<'a> ComponentOutput<'a> {
    fn new(segmented: bool) -> Self {
        let mut header = Vec::with_capacity(8);
        header.extend_from_slice(&WASM_MAGIC);
        header.extend_from_slice(&[0x0d, 0, 1, 0]);
        if segmented {
            let len = header.len();
            let mut parts = Vec::with_capacity(4);
            parts.push(Cow::Owned(header));
            Self::Segmented { len, parts }
        } else {
            Self::Contiguous(header)
        }
    }

    fn len(&self) -> usize {
        match self {
            Self::Contiguous(bytes) => bytes.len(),
            Self::Segmented { len, .. } => *len,
        }
    }

    fn push_part(&mut self, part: Cow<'a, [u8]>) -> Result<(), String> {
        let Self::Segmented { parts, len } = self else { unreachable!() };
        *len = len.checked_add(part.len()).ok_or("component binary length overflows usize")?;
        parts.push(part);
        Ok(())
    }

    fn section_header(&mut self, id: u8, len: usize) -> Result<(), String> {
        let len = checked_len(len, "component section")?;
        let mut header = Vec::with_capacity(6);
        header.push(id);
        write_leb128_u32(&mut header, len);
        self.push_part(Cow::Owned(header))
    }

    fn owned_section(&mut self, id: u8, data: &[u8]) -> Result<(), String> {
        let size = checked_len(data.len(), "component section")?;
        match self {
            Self::Contiguous(bytes) => write_section(bytes, id, data),
            Self::Segmented { parts, len } => {
                let header_len = 1 + (32 - size.leading_zeros()).max(1).div_ceil(7) as usize;
                let capacity = data.len().checked_add(header_len)
                    .ok_or("component section length overflows usize")?;
                let next_len = len.checked_add(capacity)
                    .ok_or("component binary length overflows usize")?;
                // Borrowed payloads remain barriers; adjacent metadata shares a buffer.
                if let Some(Cow::Owned(section)) = parts.last_mut() {
                    write_section(section, id, data);
                } else {
                    let mut section = Vec::with_capacity(capacity);
                    write_section(&mut section, id, data);
                    parts.push(Cow::Owned(section));
                }
                *len = next_len;
            }
        }
        Ok(())
    }

    fn borrowed_section(&mut self, id: u8, data: &'a [u8]) -> Result<(), String> {
        checked_len(data.len(), "component section")?;
        match self {
            Self::Contiguous(bytes) => write_section(bytes, id, data),
            Self::Segmented { .. } => {
                self.section_header(id, data.len())?;
                self.push_part(Cow::Borrowed(data))?;
            }
        }
        Ok(())
    }

    fn custom_section(&mut self, name: &[u8], data: &'a [u8]) -> Result<(), String> {
        let len = name.len().checked_add(data.len()).ok_or("component section length overflows usize")?;
        let len = checked_len(len, "component section")?;
        match self {
            Self::Contiguous(bytes) => {
                bytes.push(0);
                write_leb128_u32(bytes, len);
                bytes.extend_from_slice(name);
                bytes.extend_from_slice(data);
            }
            Self::Segmented { .. } => {
                let mut header = Vec::with_capacity(6 + name.len());
                header.push(0);
                write_leb128_u32(&mut header, len);
                header.extend_from_slice(name);
                self.push_part(Cow::Owned(header))?;
                self.push_part(Cow::Borrowed(data))?;
            }
        }
        Ok(())
    }

    fn nested(&mut self, child: Self) -> Result<(), String> {
        checked_len(child.len(), "component section")?;
        match (self, child) {
            (Self::Contiguous(bytes), Self::Contiguous(child)) => write_section(bytes, 4, &child),
            (parent @ Self::Segmented { .. }, Self::Segmented { mut parts, len }) => {
                let Self::Segmented { parts: parent_parts, .. } = parent else { unreachable!() };
                parent_parts.reserve(1 + parts.len());
                parent.section_header(4, len)?;
                let Self::Segmented { parts: parent_parts, len: parent_len } = parent else { unreachable!() };
                *parent_len = parent_len.checked_add(len).ok_or("component binary length overflows usize")?;
                parent_parts.append(&mut parts);
            }
            _ => unreachable!("nested components share their parent's output mode"),
        }
        Ok(())
    }

    fn flatten(self) -> Vec<u8> {
        match self {
            Self::Contiguous(bytes) => bytes,
            Self::Segmented { parts, len } => {
                let mut bytes = Vec::with_capacity(len);
                for part in parts { bytes.extend_from_slice(&part); }
                bytes
            }
        }
    }
}

fn has_bulk_payload(component: &ComponentBinary) -> bool {
    // Keep small metadata-only components on the allocation-light contiguous path.
    component.definitions.iter().any(|definition| match definition {
        ComponentBinaryDef::Custom { data, .. } | ComponentBinaryDef::CoreModule(data) => data.len() >= 64 * 1024,
        ComponentBinaryDef::NestedComponent(child) => has_bulk_payload(child),
        _ => false,
    })
}

fn write_component_section(out: &mut ComponentOutput<'_>, id: u8, data: &[u8]) -> Result<(), String> {
    out.owned_section(id, data)
}

fn write_import_func(
    out: &mut Vec<u8>,
    name: &str,
    type_index: u32,
    type_count: u32,
) -> Result<(), String> {
    if type_index >= type_count {
        return Err(format!("component type index {type_index} is not declared"));
    }
    out.push(0);
    write_component_name(out, name)?;
    out.push(1);
    write_leb128_u32(out, type_index);
    Ok(())
}

fn write_export_func(
    out: &mut Vec<u8>,
    name: &str,
    index: u32,
    type_index: Option<u32>,
    component_func_count: &mut u32,
) -> Result<(), String> {
    if index >= *component_func_count {
        return Err(format!("component function index {index} is not declared"));
    }
    out.push(0);
    write_component_name(out, name)?;
    out.push(1);
    write_leb128_u32(out, index);
    match type_index {
        Some(index) => {
            out.push(1);
            out.push(1); // function externtype
            write_leb128_u32(out, index);
        }
        None => out.push(0),
    }
    *component_func_count = component_func_count
        .checked_add(1)
        .ok_or("component function index overflow")?;
    Ok(())
}

/// Encode a Component Model binary. The accepted IR is deliberately narrower
/// than the proposal: unsupported definitions must be represented before they
/// can reach this function, rather than being dropped from the output.
pub fn write_component_binary(component: &ComponentBinary) -> Result<Vec<u8>, String> {
    write_component_binary_impl(component, true, false, true)
}

fn write_component_binary_impl(
    component: &ComponentBinary,
    validate_output: bool,
    validate_children: bool,
    group_imports_and_instances: bool,
) -> Result<Vec<u8>, String> {
    let segmented = !validate_children && component.definitions.iter().any(|definition| {
        matches!(definition, ComponentBinaryDef::NestedComponent(child) if has_bulk_payload(child))
    });
    let out = write_component_output(component, validate_output && !segmented,
        validate_children, group_imports_and_instances, segmented)?.flatten();
    if validate_output && segmented {
        crate::reader::native_validation::validate_component(&out)
            .map_err(|e| e.to_string())?;
    }
    Ok(out)
}

fn write_component_output<'a>(
    component: &'a ComponentBinary,
    validate_output: bool,
    validate_children: bool,
    group_imports_and_instances: bool,
    segmented: bool,
) -> Result<ComponentOutput<'a>, String> {
    checked_len(component.definitions.len(), "component definitions")?;
    let mut out = ComponentOutput::new(segmented);

    let mut module_count = 0_u32;
    let mut instance_count = 0_u32;
    let mut component_count = 0_u32;
    let mut component_instance_count = 0_u32;
    let mut type_count = 0_u32;
    let mut core_func_count = 0_u32;
    let mut core_memory_count = 0_u32;
    let mut core_type_count = 0_u32;
    let mut core_table_count = 0_u32;
    let mut component_func_count = 0_u32;
    let mut component_value_count = 0_u32;
    let mut section = Vec::new();
    let mut definitions = component.definitions.iter().peekable();
    while let Some(definition) = definitions.next() {
        match definition {
            ComponentBinaryDef::CoreFuncType { params, results } => {
                let mut types = Vec::new();
                write_core_function_type(&mut types, params, results)?;
                let mut count = 1_u32;
                while let Some(ComponentBinaryDef::CoreFuncType { params, results }) = definitions.peek() {
                    write_core_function_type(&mut types, params, results)?;
                    definitions.next();
                    count = count.checked_add(1).ok_or("core type section count overflow")?;
                }
                section.clear();
                write_leb128_u32(&mut section, count);
                section.extend_from_slice(&types);
                write_component_section(&mut out, 3, &section)?;
                core_type_count = core_type_count.checked_add(count)
                    .ok_or("core type index overflow")?;
            }
            ComponentBinaryDef::ValueType(_)
            | ComponentBinaryDef::FuncType { .. }
            | ComponentBinaryDef::InterfaceType { .. }
            | ComponentBinaryDef::ResourceType { .. } => {
                let mut types = Vec::new();
                write_type_definition(&mut types, definition, type_count)?;
                type_count = type_count
                    .checked_add(1)
                    .ok_or("component type index overflow")?;
                let mut count = 1_u32;
                while matches!(
                    definitions.peek(),
                    Some(
                        ComponentBinaryDef::ValueType(_)
                            | ComponentBinaryDef::FuncType { .. }
                            | ComponentBinaryDef::InterfaceType { .. }
                            | ComponentBinaryDef::ResourceType { .. }
                    )
                ) {
                    let next = definitions
                        .next()
                        .ok_or("missing component type definition")?;
                    write_type_definition(&mut types, next, type_count)?;
                    type_count = type_count
                        .checked_add(1)
                        .ok_or("component type index overflow")?;
                    count = count
                        .checked_add(1)
                        .ok_or("component type section count overflow")?;
                }
                section.clear();
                write_leb128_u32(&mut section, count);
                section.extend_from_slice(&types);
                write_component_section(&mut out, 7, &section)?;
            }
            ComponentBinaryDef::ImportFunc { .. } | ComponentBinaryDef::Import { .. } => {
                let additional = if group_imports_and_instances {
                    definitions
                        .clone()
                        .take_while(|definition| {
                            matches!(definition, ComponentBinaryDef::ImportFunc { .. } | ComponentBinaryDef::Import { .. })
                        })
                        .count()
                } else {
                    0
                };
                section.clear();
                write_leb128_u32(
                    &mut section,
                    checked_len(additional + 1, "component imports")?,
                );
                let mut import = definition;
                for position in 0..=additional {
                    if position != 0 {
                        import = definitions.next().ok_or("missing component import")?;
                    }
                    let increment = match import {
                        ComponentBinaryDef::ImportFunc { name, type_index } => {
                            write_import_func(&mut section, name, *type_index, type_count)?;
                            Some(&mut component_func_count)
                        }
                        ComponentBinaryDef::Import { name, ty } => {
                            section.push(0);
                            write_component_name(&mut section, name)?;
                            write_interface_ref(&mut section, ty, type_count)?;
                            match ty {
                                ComponentInterfaceRef::Module(_) => Some(&mut module_count),
                                ComponentInterfaceRef::Func(_) => Some(&mut component_func_count),
                                ComponentInterfaceRef::TypeEq(_) | ComponentInterfaceRef::TypeSubResource => Some(&mut type_count),
                                ComponentInterfaceRef::Component(_) => Some(&mut component_count),
                                ComponentInterfaceRef::Instance(_) => Some(&mut component_instance_count),
                                ComponentInterfaceRef::Value(_) => Some(&mut component_value_count),
                            }
                        }
                        _ => return Err("expected adjacent component import".into()),
                    };
                    if let Some(count) = increment {
                        *count = count
                        .checked_add(1)
                        .ok_or("component import index overflow")?;
                    }
                }
                write_component_section(&mut out, 10, &section)?;
            }
            ComponentBinaryDef::ExportFunc { name, index, type_index } => {
                let additional = definitions.clone()
                    .take_while(|definition| matches!(definition, ComponentBinaryDef::ExportFunc { .. }))
                    .count();
                section.clear();
                write_leb128_u32(&mut section, checked_len(additional + 1, "component exports")?);
                write_export_func(&mut section, name, *index, *type_index, &mut component_func_count)?;
                for _ in 0..additional {
                    let Some(ComponentBinaryDef::ExportFunc { name, index, type_index }) = definitions.next()
                    else { unreachable!("counted only adjacent function exports") };
                    write_export_func(&mut section, name, *index, *type_index, &mut component_func_count)?;
                }
                write_component_section(&mut out, 11, &section)?;
            }
            ComponentBinaryDef::NestedComponent(nested) => {
                let child = write_component_output(
                    nested,
                    validate_children,
                    validate_children,
                    group_imports_and_instances,
                    segmented,
                )?;
                out.nested(child)?;
                component_count = component_count
                    .checked_add(1)
                    .ok_or("component index overflow")?;
            }
            ComponentBinaryDef::ComponentInstance { .. } | ComponentBinaryDef::InstanceFromExports(_) => {
                let additional = if group_imports_and_instances {
                    definitions.clone().take_while(|definition| matches!(definition,
                        ComponentBinaryDef::ComponentInstance { .. } | ComponentBinaryDef::InstanceFromExports(_))).count()
                } else { 0 };
                section.clear();
                write_leb128_u32(&mut section, checked_len(additional + 1, "component instances")?);
                let mut instance = definition;
                for position in 0..=additional {
                    if position != 0 {
                        instance = definitions.next().ok_or("missing component instance")?;
                    }
                    let counts = [module_count, component_func_count, component_value_count,
                        type_count, component_count, component_instance_count];
                    match instance {
                        ComponentBinaryDef::ComponentInstance { component, args } => {
                            if *component >= component_count {
                                return Err(format!("component index {component} is not declared"));
                            }
                            section.push(0);
                            write_leb128_u32(&mut section, *component);
                            write_leb128_u32(&mut section, checked_len(args.len(), "component arguments")?);
                            for arg in args {
                                write_component_name(&mut section, &arg.name)?;
                                write_item_ref(&mut section, arg, counts)?;
                            }
                        }
                        ComponentBinaryDef::InstanceFromExports(exports) => {
                            section.push(1);
                            write_leb128_u32(&mut section, checked_len(exports.len(), "component instance exports")?);
                            for export in exports {
                                section.push(0);
                                write_component_name(&mut section, &export.name)?;
                                write_item_ref(&mut section, export, counts)?;
                            }
                        }
                        _ => return Err("expected adjacent component instance".into()),
                    }
                    component_instance_count = component_instance_count.checked_add(1)
                        .ok_or("component instance index overflow")?;
                }
                write_component_section(&mut out, 5, &section)?;
            }
            ComponentBinaryDef::AliasInstanceExport { .. } => {
                let additional = definitions.clone().take_while(|definition|
                    matches!(definition, ComponentBinaryDef::AliasInstanceExport { .. })).count();
                section.clear();
                write_leb128_u32(&mut section, checked_len(additional + 1, "component aliases")?);
                let mut alias = definition;
                for position in 0..=additional {
                    if position != 0 {
                        alias = definitions.next().ok_or("missing component alias")?;
                    }
                    let ComponentBinaryDef::AliasInstanceExport { instance, name, kind } = alias
                    else { unreachable!("counted only adjacent instance-export aliases") };
                    if *instance >= component_instance_count {
                        return Err(format!("component instance index {instance} is not declared"));
                    }
                    write_item_sort(&mut section, *kind);
                    section.push(0);
                    write_leb128_u32(&mut section, *instance);
                    write_component_name(&mut section, name)?;
                    use vybe_ast::canon::ComponentItemKind as K;
                    let count = match kind {
                        K::Module => &mut module_count,
                        K::Func => &mut component_func_count,
                        K::Value => &mut component_value_count,
                        K::Type => &mut type_count,
                        K::Component => &mut component_count,
                        K::Instance => &mut component_instance_count,
                    };
                    *count = count.checked_add(1).ok_or("component alias index overflow")?;
                }
                write_component_section(&mut out, 6, &section)?;
            }
            ComponentBinaryDef::ExportInstance { name, index } => {
                if *index >= component_instance_count {
                    return Err("instance index out of bounds".into());
                }
                section.clear();
                section.extend_from_slice(&[1, 0]);
                write_component_name(&mut section, name)?;
                section.push(5);
                write_leb128_u32(&mut section, *index);
                section.push(0);
                write_component_section(&mut out, 11, &section)?;
                component_instance_count = component_instance_count.checked_add(1)
                    .ok_or("component instance index overflow")?;
            }
            ComponentBinaryDef::ExportCoreModule { name, index } => {
                if *index >= module_count {
                    return Err("module index out of bounds".into());
                }
                section.clear();
                section.extend_from_slice(&[1, 0]);
                write_component_name(&mut section, name)?;
                section.extend_from_slice(&[0, 0x11]);
                write_leb128_u32(&mut section, *index);
                section.push(0);
                write_component_section(&mut out, 11, &section)?;
                module_count = module_count.checked_add(1)
                    .ok_or("core module index overflow")?;
            }
            ComponentBinaryDef::ExportType { name, index } => {
                if *index >= type_count {
                    return Err(format!("component type index {index} is not declared"));
                }
                section.clear();
                section.push(1);
                section.push(0);
                write_component_name(&mut section, name)?;
                section.push(3);
                write_leb128_u32(&mut section, *index);
                section.push(0);
                write_component_section(&mut out, 11, &section)?;
                type_count = type_count
                    .checked_add(1)
                    .ok_or("component type index overflow")?;
            }
            ComponentBinaryDef::CanonLift { .. }
            | ComponentBinaryDef::CanonLower { .. }
            | ComponentBinaryDef::CanonTaskReturn { .. }
            | ComponentBinaryDef::CanonThreadIndex
            | ComponentBinaryDef::CanonResourceNew { .. }
            | ComponentBinaryDef::CanonResourceDrop { .. }
            | ComponentBinaryDef::CanonResourceRep { .. }
            | ComponentBinaryDef::CanonErrorContextNew { .. }
            | ComponentBinaryDef::CanonErrorContextDebugMessage { .. }
            | ComponentBinaryDef::CanonErrorContextDrop
            | ComponentBinaryDef::CanonWaitableSetNew
            | ComponentBinaryDef::CanonWaitableSetWait { .. }
            | ComponentBinaryDef::CanonWaitableSetPoll { .. }
            | ComponentBinaryDef::CanonWaitableSetDrop
            | ComponentBinaryDef::CanonWaitableJoin
            | ComponentBinaryDef::CanonContext { .. }
            | ComponentBinaryDef::CanonTaskCancel
            | ComponentBinaryDef::CanonSubtaskCancel { .. }
            | ComponentBinaryDef::CanonSubtaskDrop
            | ComponentBinaryDef::CanonBackpressureInc
            | ComponentBinaryDef::CanonBackpressureDec
            | ComponentBinaryDef::CanonThreadResumeLater
            | ComponentBinaryDef::CanonThreadSuspend
            | ComponentBinaryDef::CanonThreadYield
            | ComponentBinaryDef::CanonThreadSuspendThenResume
            | ComponentBinaryDef::CanonThreadYieldThenResume
            | ComponentBinaryDef::CanonThreadSuspendThenPromote
            | ComponentBinaryDef::CanonThreadYieldThenPromote
            | ComponentBinaryDef::CanonThreadNewIndirect { .. }
            | ComponentBinaryDef::CanonAsync { .. } => {
                let mut canons = Vec::new();
                write_canon_definition(
                    &mut canons,
                    definition,
                    type_count,
                    &mut core_func_count,
                    &mut component_func_count,
                    core_memory_count,
                    core_type_count,
                    core_table_count,
                )?;
                let mut count = 1_u32;
                while matches!(
                    definitions.peek(),
                    Some(
                        ComponentBinaryDef::CanonLift { .. }
                            | ComponentBinaryDef::CanonLower { .. }
                            | ComponentBinaryDef::CanonTaskReturn { .. }
                            | ComponentBinaryDef::CanonThreadIndex
                            | ComponentBinaryDef::CanonResourceNew { .. }
                            | ComponentBinaryDef::CanonResourceDrop { .. }
                            | ComponentBinaryDef::CanonResourceRep { .. }
                            | ComponentBinaryDef::CanonErrorContextNew { .. }
                            | ComponentBinaryDef::CanonErrorContextDebugMessage { .. }
                            | ComponentBinaryDef::CanonErrorContextDrop
                            | ComponentBinaryDef::CanonWaitableSetNew
                            | ComponentBinaryDef::CanonWaitableSetWait { .. }
                            | ComponentBinaryDef::CanonWaitableSetPoll { .. }
                            | ComponentBinaryDef::CanonWaitableSetDrop
                            | ComponentBinaryDef::CanonWaitableJoin
                            | ComponentBinaryDef::CanonContext { .. }
                            | ComponentBinaryDef::CanonTaskCancel
                            | ComponentBinaryDef::CanonSubtaskCancel { .. }
                            | ComponentBinaryDef::CanonSubtaskDrop
                            | ComponentBinaryDef::CanonBackpressureInc
                            | ComponentBinaryDef::CanonBackpressureDec
                            | ComponentBinaryDef::CanonThreadResumeLater
                            | ComponentBinaryDef::CanonThreadSuspend
                            | ComponentBinaryDef::CanonThreadYield
                            | ComponentBinaryDef::CanonThreadSuspendThenResume
                            | ComponentBinaryDef::CanonThreadYieldThenResume
                            | ComponentBinaryDef::CanonThreadSuspendThenPromote
                            | ComponentBinaryDef::CanonThreadYieldThenPromote
                            | ComponentBinaryDef::CanonThreadNewIndirect { .. }
                            | ComponentBinaryDef::CanonAsync { .. }
                    )
                ) {
                    let next = definitions.next().ok_or("missing canonical definition")?;
                    write_canon_definition(
                        &mut canons,
                        next,
                        type_count,
                        &mut core_func_count,
                        &mut component_func_count,
                        core_memory_count,
                        core_type_count,
                        core_table_count,
                    )?;
                    count = count
                        .checked_add(1)
                        .ok_or("canonical section count overflow")?;
                }
                section.clear();
                write_leb128_u32(&mut section, count);
                section.extend_from_slice(&canons);
                write_component_section(&mut out, 8, &section)?;
            }
            ComponentBinaryDef::Custom { name, data } => {
                section.clear();
                write_component_name(&mut section, name)?;
                out.custom_section(&section, data)?;
            }
            ComponentBinaryDef::CoreModule(bytes) => {
                if !bytes.starts_with(&[0, 0x61, 0x73, 0x6d, 1, 0, 0, 0]) {
                    return Err("core module has no standard Wasm header".into());
                }
                out.borrowed_section(1, bytes)?;
                module_count += 1;
            }
            ComponentBinaryDef::CoreInstance { module, args } => {
                if *module >= module_count {
                    return Err(format!("core module index {module} is not declared"));
                }
                section.clear();
                write_leb128_u32(&mut section, 1);
                section.push(0);
                write_leb128_u32(&mut section, *module);
                write_leb128_u32(
                    &mut section,
                    checked_len(args.len(), "core instance arguments")?,
                );
                for (name, instance) in args {
                    if *instance >= instance_count {
                        return Err(format!("core instance index {instance} is not declared"));
                    }
                    write_component_name(&mut section, name)?;
                    section.push(0x12);
                    write_leb128_u32(&mut section, *instance);
                }
                write_component_section(&mut out, 2, &section)?;
                instance_count += 1;
            }
            ComponentBinaryDef::CoreInstanceFromExports(exports) => {
                section.clear();
                write_leb128_u32(&mut section, 1);
                section.push(1);
                write_leb128_u32(
                    &mut section,
                    checked_len(exports.len(), "core instance exports")?,
                );
                for export in exports {
                    write_component_name(&mut section, &export.name)?;
                    section.push(core_kind_byte(export.kind));
                    write_leb128_u32(&mut section, export.index);
                }
                write_component_section(&mut out, 2, &section)?;
                instance_count += 1;
            }
            ComponentBinaryDef::CoreAliasInstanceExport {
                instance,
                name,
                kind,
            } => {
                if *instance >= instance_count {
                    return Err(format!("core instance index {instance} is not declared"));
                }
                section.clear();
                write_leb128_u32(&mut section, 1);
                section.push(0);
                section.push(core_kind_byte(*kind));
                section.push(1);
                write_leb128_u32(&mut section, *instance);
                write_component_name(&mut section, name)?;
                write_component_section(&mut out, 6, &section)?;
                match kind {
                    CoreExportKind::Func => {
                        core_func_count = core_func_count
                            .checked_add(1)
                            .ok_or("core function index overflow")?
                    }
                    CoreExportKind::Memory => {
                        core_memory_count = core_memory_count
                            .checked_add(1)
                            .ok_or("core memory index overflow")?
                    }
                    CoreExportKind::Table => {
                        core_table_count = core_table_count.checked_add(1)
                            .ok_or("core table index overflow")?
                    }
                    _ => {}
                }
            }
        }
    }
    if validate_output {
        let ComponentOutput::Contiguous(bytes) = &out else {
            unreachable!("segmented output is validated after flattening")
        };
        crate::reader::native_validation::validate_component(bytes)
            .map_err(|e| e.to_string())?;
    }
    Ok(out)
}

#[cfg(test)]
mod tests {
    use super::*;
    use vybe_ast::canon::CoreInstanceExport;

    #[test]
    fn segmented_metadata_sections_match_contiguous_at_length_boundaries() {
        for len in [0, 1, 127, 128, 16_383, 16_384, 65_536] {
            let data = vec![0x5a; len];
            let mut contiguous = ComponentOutput::new(false);
            let mut segmented = ComponentOutput::new(true);
            contiguous.owned_section(7, &data).unwrap();
            segmented.owned_section(7, &data).unwrap();
            assert!(matches!(&segmented, ComponentOutput::Segmented { parts, .. }
                if parts.len() == 1));
            assert_eq!(segmented.flatten(), contiguous.flatten(), "length {len}");
        }
    }

    #[test]
    fn coalesced_metadata_preserves_borrowed_payloads_and_section_order() {
        let core = module_with_empty_functions(20_000);
        let custom = [0, 42];
        let mut contiguous = ComponentOutput::new(false);
        let mut segmented = ComponentOutput::new(true);
        for output in [&mut contiguous, &mut segmented] {
            output.borrowed_section(1, &core).unwrap();
            output.owned_section(2, &[1, 1, 0]).unwrap();
            output.owned_section(2, &[1, 1, 0]).unwrap();
            output.borrowed_section(0, &custom).unwrap();
            output.owned_section(2, &[1, 1, 0]).unwrap();
        }
        let ComponentOutput::Segmented { parts, .. } = &segmented else { unreachable!() };
        let borrowed: Vec<_> = parts.iter().filter_map(|part| match part {
            Cow::Borrowed(bytes) => Some(*bytes),
            Cow::Owned(_) => None,
        }).collect();
        assert_eq!(borrowed.len(), 2);
        assert_eq!(borrowed[0].as_ptr(), core.as_ptr());
        assert_eq!(borrowed[1].as_ptr(), custom.as_ptr());
        assert_eq!(segmented.len(), contiguous.len());
        let binary = segmented.flatten();
        assert_eq!(binary, contiguous.flatten());
        wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
            .validate_all(&binary).unwrap();
    }

    #[test]
    fn coalesced_metadata_overflow_does_not_mutate_output() {
        let original = vec![1, 2, 3];
        let mut output = ComponentOutput::Segmented {
            parts: vec![Cow::Owned(original.clone())], len: usize::MAX,
        };
        assert!(output.owned_section(2, &[1, 1, 0]).unwrap_err().contains("overflows"));
        let ComponentOutput::Segmented { parts, len } = output else { unreachable!() };
        assert_eq!(len, usize::MAX);
        assert_eq!(parts.len(), 1);
        assert_eq!(parts[0].as_ref(), original);
    }

    #[test]
    #[ignore = "paired segmented metadata assembly benchmark, not overall CM latency"]
    fn benchmark_coalesced_component_metadata() {
        const SECTIONS: usize = 4_000;
        const RUNS: usize = 11;
        let core = module_with_empty_functions(20_000);
        let assemble = |coalesced: bool| {
            let mut output = ComponentOutput::new(true);
            output.borrowed_section(1, &core).unwrap();
            for _ in 0..SECTIONS {
                if coalesced {
                    output.owned_section(2, &[1, 1, 0]).unwrap();
                } else {
                    // Exact previous owned_section path, retained only for comparison.
                    let mut section = Vec::with_capacity(3 + 6);
                    write_section(&mut section, 2, &[1, 1, 0]);
                    output.push_part(Cow::Owned(section)).unwrap();
                }
            }
            let ComponentOutput::Segmented { parts, .. } = &output else { unreachable!() };
            let parts = parts.len();
            (output.flatten(), parts)
        };
        let (reference, old_parts) = assemble(false);
        let (expected, new_parts) = assemble(true);
        assert_eq!(reference, expected);
        wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
            .validate_all(&expected).unwrap();
        let mut times = [Vec::new(), Vec::new()];
        for run in 0..RUNS {
            for mode in [run % 2, 1 - run % 2] {
                let start = std::time::Instant::now();
                let (binary, _) = assemble(mode == 1);
                times[mode].push(start.elapsed());
                assert_eq!(binary, expected);
                std::hint::black_box(binary);
            }
        }
        for samples in &mut times { samples.sort_unstable(); }
        eprintln!("segmented metadata {SECTIONS} sections: {} vs {} parts; {:.3} vs {:.3} ms",
            old_parts, new_parts, times[0][RUNS / 2].as_secs_f64() * 1000.0,
            times[1][RUNS / 2].as_secs_f64() * 1000.0);
    }

    #[test]
    fn grouped_exports_preserve_alias_indices_and_definition_boundaries() {
        for count in [1, 127, 128, 129] {
            let mut definitions = vec![
                ComponentBinaryDef::FuncType { params: vec![], result: None, is_async: false },
                ComponentBinaryDef::ImportFunc { name: "f".into(), type_index: 0 },
                ComponentBinaryDef::ExportFunc { name: "first".into(), index: 0, type_index: None },
                ComponentBinaryDef::Custom { name: "separator".into(), data: vec![42] },
            ];
            definitions.extend((0..count).map(|index| ComponentBinaryDef::ExportFunc {
                name: format!("next-{index}"), index: index + 1, type_index: None,
            }));
            let component = ComponentBinary { definitions };
            let bytes = write_component_binary(&component).unwrap();
            assert_eq!(crate::read_component_binary(&bytes).unwrap(), component);
            let counts: Vec<_> = wasmparser::Parser::new(0).parse_all(&bytes)
                .filter_map(|payload| match payload.unwrap() {
                    wasmparser::Payload::ComponentExportSection(reader) => Some(reader.count()),
                    _ => None,
                }).collect();
            assert_eq!(counts, [1, count]);
        }
        let invalid = ComponentBinary { definitions: vec![
            ComponentBinaryDef::FuncType { params: vec![], result: None, is_async: false },
            ComponentBinaryDef::ImportFunc { name: "f".into(), type_index: 0 },
            ComponentBinaryDef::ExportFunc { name: "first".into(), index: 0, type_index: None },
            ComponentBinaryDef::ExportFunc { name: "forward".into(), index: 2, type_index: None },
        ] };
        assert!(write_component_binary(&invalid).unwrap_err()
            .contains("component function index 2 is not declared"));
    }

    #[test]
    fn function_export_ascriptions_roundtrip_and_validate_compatibility() {
        let mut component = ComponentBinary { definitions: vec![
            ComponentBinaryDef::FuncType { params: vec![], result: None, is_async: false },
            ComponentBinaryDef::FuncType { params: vec![], result: None, is_async: false },
            ComponentBinaryDef::ImportFunc { name: "f".into(), type_index: 0 },
            ComponentBinaryDef::ExportFunc { name: "ascribed".into(), index: 0, type_index: Some(1) },
            ComponentBinaryDef::ExportFunc { name: "alias".into(), index: 1, type_index: Some(0) },
        ] };
        let bytes = write_component_binary(&component).unwrap();
        assert_eq!(crate::read_component_binary(&bytes).unwrap(), component);
        component.definitions[1] = ComponentBinaryDef::FuncType {
            params: vec![("extra".into(), ValSpec::Prim("u32".into()))],
            result: None, is_async: false,
        };
        assert!(write_component_binary(&component).unwrap_err()
            .contains("ascribed type of export is not compatible"));
        component.definitions[1] = ComponentBinaryDef::ValueType(ValSpec::Prim("u32".into()));
        assert!(write_component_binary(&component).unwrap_err()
            .contains("not a function type"));
    }

    #[test]
    fn canonical_stream_future_operations_round_trip() {
        use ComponentBinaryAsyncOperation::*;
        let memory = vec![0, 97, 115, 109, 1, 0, 0, 0,
            5, 3, 1, 0, 1, 7, 7, 1, 3, b'm', b'e', b'm', 2, 0];
        let mut component = ComponentBinary { definitions: vec![
            ComponentBinaryDef::CoreModule(memory),
            ComponentBinaryDef::CoreInstance { module: 0, args: vec![] },
            ComponentBinaryDef::CoreAliasInstanceExport {
                instance: 0, name: "mem".into(), kind: CoreExportKind::Memory,
            },
            ComponentBinaryDef::ValueType(ValSpec::Stream(Some(Box::new(ValSpec::Prim("u8".into()))))),
            ComponentBinaryDef::ValueType(ValSpec::Future(Some(Box::new(ValSpec::Prim("u8".into()))))),
        ] };
        for (type_index, kind) in [ComponentBinaryAsyncKind::Stream,
            ComponentBinaryAsyncKind::Future].into_iter().enumerate()
        {
            for operation in [New, Read, Write, CancelRead { is_async: false },
                CancelRead { is_async: true }, CancelWrite { is_async: false },
                CancelWrite { is_async: true }, DropReadable, DropWritable, Forward]
            {
                component.definitions.push(ComponentBinaryDef::CanonAsync {
                    kind, operation, type_index: type_index as u32,
                    options: if matches!(operation, Read | Write) {
                        vec![ComponentBinaryCanonOption::Memory(0)]
                    } else { vec![] },
                });
            }
        }
        let bytes = write_component_binary(&component).expect("standard canonical definitions");
        assert_eq!(crate::read_component_binary(&bytes).unwrap(), component);
    }

    #[test]
    fn canonical_async_definitions_reject_undeclared_types_and_extra_options() {
        for (type_index, options, expected) in [
            (1, vec![], "component type index 1 is not declared"),
            (0, vec![ComponentBinaryCanonOption::Async], "does not accept options"),
        ] {
            let component = ComponentBinary { definitions: vec![
                ComponentBinaryDef::ValueType(ValSpec::Stream(None)),
                ComponentBinaryDef::CanonAsync {
                    kind: ComponentBinaryAsyncKind::Stream,
                    operation: ComponentBinaryAsyncOperation::New, type_index, options,
                },
            ] };
            assert!(write_component_binary(&component).unwrap_err().contains(expected));
        }
    }

    #[test]
    fn instance_exports_append_aliases_without_reinstantiation() {
        let component = ComponentBinary { definitions: vec![
            ComponentBinaryDef::NestedComponent(std::sync::Arc::new(ComponentBinary { definitions: vec![] })),
            ComponentBinaryDef::ComponentInstance { component: 0, args: vec![] },
            ComponentBinaryDef::ExportInstance { name: "first".into(), index: 0 },
            ComponentBinaryDef::ExportInstance { name: "second".into(), index: 1 },
            ComponentBinaryDef::ExportInstance { name: "third".into(), index: 2 },
        ] };
        let bytes = write_component_binary(&component).expect("instance export aliases validate");
        assert_eq!(crate::read_component_binary(&bytes).unwrap(), component);
        let invalid = ComponentBinary { definitions: vec![
            ComponentBinaryDef::ExportInstance { name: "absent".into(), index: 0 },
        ] };
        assert!(write_component_binary(&invalid).unwrap_err().contains("instance index out of bounds"));
    }

    #[test]
    fn core_module_exports_append_to_the_module_index_space() {
        let component = ComponentBinary { definitions: vec![
            ComponentBinaryDef::CoreModule(vec![0, 0x61, 0x73, 0x6d, 1, 0, 0, 0]),
            ComponentBinaryDef::ExportCoreModule { name: "first".into(), index: 0 },
            ComponentBinaryDef::ExportCoreModule { name: "second".into(), index: 1 },
            ComponentBinaryDef::CoreInstance { module: 2, args: vec![] },
        ] };
        let bytes = write_component_binary(&component).expect("export aliases validate");
        assert_eq!(crate::read_component_binary(&bytes).unwrap(), component);
        let invalid = ComponentBinary { definitions: vec![
            ComponentBinaryDef::ExportCoreModule { name: "absent".into(), index: 0 },
        ] };
        assert!(write_component_binary(&invalid).unwrap_err().contains("module index out of bounds"));
    }

    #[test]
    fn primitive_type_export_matches_component_reference() {
        let component = ComponentBinary {
            definitions: vec![
                ComponentBinaryDef::ValueType(ValSpec::Prim("u32".into())),
                ComponentBinaryDef::ExportType {
                    name: "t".into(),
                    index: 0,
                },
            ],
        };
        let bytes = write_component_binary(&component).expect("standard component binary");
        assert_eq!(
            bytes,
            vec![
                0, 0x61, 0x73, 0x6d, 0x0d, 0, 1, 0, 7, 2, 1, 0x79, 11, 7, 1, 0, 1, b't', 3, 0, 0,
            ]
        );
        assert_eq!(crate::read_component_binary(&bytes).unwrap(), component);
    }

    #[test]
    fn current_thread_protocol_round_trips_with_reserved_zero_bytes() {
        let component = ComponentBinary { definitions: vec![
            ComponentBinaryDef::CanonThreadResumeLater,
            ComponentBinaryDef::CanonThreadSuspend,
            ComponentBinaryDef::CanonThreadYield,
            ComponentBinaryDef::CanonThreadSuspendThenResume,
            ComponentBinaryDef::CanonThreadYieldThenResume,
            ComponentBinaryDef::CanonThreadSuspendThenPromote,
            ComponentBinaryDef::CanonThreadYieldThenPromote,
        ] };
        let bytes = write_component_binary(&component).expect("current CM thread protocol");
        assert_eq!(&bytes[8..], &[8, 14, 7, 0x28, 0x29, 0, 0x0c, 0,
            0x2a, 0, 0x2b, 0, 0x2c, 0, 0x2d, 0]);
        assert_eq!(crate::read_component_binary(&bytes).unwrap(), component);
        for offset in [13, 15, 17, 19, 21, 23] {
            let mut invalid = bytes.clone();
            invalid[offset] = 1;
            assert!(crate::read_component_binary(&invalid).is_err());
        }
    }

    #[test]
    fn error_context_options_round_trip_and_require_sync_memory_and_realloc() {
        let module = vec![
            0, 97, 115, 109, 1, 0, 0, 0,
            1, 9, 1, 0x60, 4, 0x7f, 0x7f, 0x7f, 0x7f, 1, 0x7f,
            3, 2, 1, 0, 5, 3, 1, 0, 1,
            7, 15, 2, 3, b'm', b'e', b'm', 2, 0, 5, b'a', b'l', b'l', b'o', b'c', 0, 0,
            10, 6, 1, 4, 0, 0x41, 32, 11,
        ];
        let mut component = ComponentBinary { definitions: vec![
            ComponentBinaryDef::CoreModule(module),
            ComponentBinaryDef::CoreInstance { module: 0, args: vec![] },
            ComponentBinaryDef::CoreAliasInstanceExport { instance: 0, name: "mem".into(), kind: CoreExportKind::Memory },
            ComponentBinaryDef::CoreAliasInstanceExport { instance: 0, name: "alloc".into(), kind: CoreExportKind::Func },
            ComponentBinaryDef::CanonErrorContextNew { options: vec![ComponentBinaryCanonOption::Memory(0), ComponentBinaryCanonOption::Utf16] },
            ComponentBinaryDef::CanonErrorContextDebugMessage { options: vec![ComponentBinaryCanonOption::Memory(0), ComponentBinaryCanonOption::Realloc(0), ComponentBinaryCanonOption::Utf8] },
            ComponentBinaryDef::CanonErrorContextDrop,
        ] };
        let bytes = write_component_binary(&component).unwrap();
        assert_eq!(crate::read_component_binary(&bytes).unwrap(), component);
        let range = wasmparser::Parser::new(0).parse_all(&bytes).find_map(|payload| match payload.unwrap() {
            wasmparser::Payload::ComponentCanonicalSection(reader) => Some(reader.range()),
            _ => None,
        }).unwrap();
        assert_eq!(&bytes[range.start as usize..range.end as usize],
            &[3, 0x1c, 2, 3, 0, 1, 0x1d, 3, 3, 0, 4, 0, 0, 0x1e]);
        for options in [vec![], vec![ComponentBinaryCanonOption::Memory(1)],
            vec![ComponentBinaryCanonOption::Memory(0), ComponentBinaryCanonOption::Async]] {
            component.definitions[4] = ComponentBinaryDef::CanonErrorContextNew { options };
            assert!(write_component_binary(&component).is_err());
        }
        component.definitions[4] = ComponentBinaryDef::CanonErrorContextNew { options: vec![ComponentBinaryCanonOption::Memory(0)] };
        component.definitions[5] = ComponentBinaryDef::CanonErrorContextDebugMessage { options: vec![ComponentBinaryCanonOption::Memory(0)] };
        assert!(write_component_binary(&component).is_err());
    }

    #[test]
    fn resource_operations_match_standard_bytes_and_validate_type_identity() {
        let mut component = ComponentBinary { definitions: vec![
            ComponentBinaryDef::ResourceType { rep: ResourceRep::I32, dtor: None },
            ComponentBinaryDef::CanonResourceNew { type_index: 0 },
            ComponentBinaryDef::CanonResourceDrop { type_index: 0 },
            ComponentBinaryDef::CanonResourceRep { type_index: 0 },
        ] };
        let bytes = write_component_binary(&component).unwrap();
        assert_eq!(&bytes[8..], &[7, 4, 1, 0x3f, 0x7f, 0, 8, 7, 3, 2, 0, 3, 0, 4, 0]);
        assert_eq!(crate::read_component_binary(&bytes).unwrap(), component);
        let mut invalid = bytes.clone();
        *invalid.last_mut().unwrap() = 1;
        assert!(crate::read_component_binary(&invalid).is_err());
        component.definitions[1] = ComponentBinaryDef::CanonResourceNew { type_index: 1 };
        assert!(write_component_binary(&component).unwrap_err().contains("not declared"));
        component.definitions[1] = ComponentBinaryDef::CanonResourceNew { type_index: 0 };
        component.definitions[0] = ComponentBinaryDef::ValueType(ValSpec::Prim("u32".into()));
        assert!(write_component_binary(&component).is_err());
    }

    #[test]
    fn waitable_definitions_preserve_memory_and_reserved_zero_bytes() {
        let memory = vec![0, 97, 115, 109, 1, 0, 0, 0,
            5, 3, 1, 0, 1, 7, 7, 1, 3, b'm', b'e', b'm', 2, 0];
        let mut component = ComponentBinary { definitions: vec![
            ComponentBinaryDef::CoreModule(memory),
            ComponentBinaryDef::CoreInstance { module: 0, args: vec![] },
            ComponentBinaryDef::CoreAliasInstanceExport {
                instance: 0, name: "mem".into(), kind: CoreExportKind::Memory,
            },
            ComponentBinaryDef::CanonWaitableSetNew,
            ComponentBinaryDef::CanonWaitableSetWait { memory: 0 },
            ComponentBinaryDef::CanonWaitableSetPoll { memory: 0 },
            ComponentBinaryDef::CanonWaitableSetDrop,
            ComponentBinaryDef::CanonWaitableJoin,
        ] };
        let bytes = write_component_binary(&component).unwrap();
        assert_eq!(crate::read_component_binary(&bytes).unwrap(), component);
        let range = wasmparser::Parser::new(0).parse_all(&bytes).find_map(|p| match p.unwrap() {
            wasmparser::Payload::ComponentCanonicalSection(r) => Some(r.range()),
            _ => None,
        }).unwrap();
        let start = usize::try_from(range.start).unwrap();
        let end = usize::try_from(range.end).unwrap();
        assert_eq!(&bytes[start..end], &[5, 0x1f, 0x20, 0, 0, 0x21, 0, 0, 0x22, 0x23]);
        for offset in [3, 6] {
            let mut invalid = bytes.clone();
            invalid[start + offset] = 1;
            assert!(crate::read_component_binary(&invalid).is_err());
        }
        component.definitions.push(ComponentBinaryDef::CanonWaitableSetWait { memory: 1 });
        assert!(write_component_binary(&component).unwrap_err().contains("core memory index 1 is not declared"));
    }

    #[test]
    fn context_and_task_control_round_trip_and_reject_invalid_contexts() {
        use vybe_ast::canon::ComponentCoreValType as T;
        for ty in [T::I32, T::I64] {
            let component = ComponentBinary { definitions: vec![
                ComponentBinaryDef::CanonContext { set: false, ty, slot: 0 },
                ComponentBinaryDef::CanonContext { set: true, ty, slot: 1 },
                ComponentBinaryDef::CanonTaskCancel,
                ComponentBinaryDef::CanonSubtaskCancel { is_async: false },
                ComponentBinaryDef::CanonSubtaskCancel { is_async: true },
                ComponentBinaryDef::CanonSubtaskDrop,
                ComponentBinaryDef::CanonBackpressureInc,
                ComponentBinaryDef::CanonBackpressureDec,
            ] };
            let bytes = write_component_binary(&component).unwrap();
            assert_eq!(crate::read_component_binary(&bytes).unwrap(), component);
        }
        for (ty, slot) in [(T::F32, 0), (T::I32, 2)] {
            let invalid = ComponentBinary { definitions: vec![
                ComponentBinaryDef::CanonContext { set: false, ty, slot },
            ] };
            assert!(write_component_binary(&invalid).is_err());
        }
        let mixed = ComponentBinary { definitions: vec![
            ComponentBinaryDef::CanonContext { set: false, ty: T::I32, slot: 0 },
            ComponentBinaryDef::CanonContext { set: true, ty: T::I64, slot: 1 },
        ] };
        assert!(write_component_binary(&mixed).is_err());
    }

    #[test]
    fn core_numeric_function_types_round_trip() {
        use vybe_ast::canon::ComponentCoreValType as T;
        let component = ComponentBinary { definitions: vec![
            ComponentBinaryDef::CoreFuncType {
                params: vec![T::I32], results: vec![],
            },
            ComponentBinaryDef::CoreFuncType {
                params: vec![T::I32, T::I64, T::F32, T::F64, T::V128],
                results: vec![T::I32, T::I64, T::F32, T::F64, T::V128],
            },
            ComponentBinaryDef::CoreFuncType { params: vec![], results: vec![] },
        ] };
        let bytes = write_component_binary(&component).expect("valid core worker types");
        assert_eq!(crate::read_component_binary(&bytes).unwrap(), component);
        let sections = wasmparser::Parser::new(0).parse_all(&bytes)
            .filter_map(|payload| match payload.unwrap() {
                wasmparser::Payload::CoreTypeSection(reader) => Some(reader.count()),
                _ => None,
            }).collect::<Vec<_>>();
        assert_eq!(sections, vec![3], "adjacent types share one section");
    }

    #[test]
    fn grouped_type_counts_preserve_leb_boundary_and_prior_indices() {
        let mut definitions = vec![ComponentBinaryDef::ValueType(ValSpec::Prim("u32".into())); 128];
        definitions.push(ComponentBinaryDef::Custom { name: "split".into(), data: vec![] });
        definitions.push(ComponentBinaryDef::ValueType(ValSpec::List(Box::new(ValSpec::Ref(127)))));
        let component = ComponentBinary { definitions };
        let bytes = write_component_binary(&component).unwrap();
        assert_eq!(crate::read_component_binary(&bytes).unwrap(), component);
        let counts: Vec<_> = wasmparser::Parser::new(0).parse_all(&bytes)
            .filter_map(|payload| match payload.unwrap() {
                wasmparser::Payload::ComponentTypeSection(section) => Some(section.count()),
                _ => None,
            }).collect();
        assert_eq!(counts, [128, 1]);

        let invalid = ComponentBinary { definitions: vec![
            ComponentBinaryDef::ValueType(ValSpec::List(Box::new(ValSpec::Ref(0)))),
        ] };
        assert!(write_component_binary(&invalid).is_err());
    }

    #[test]
    fn core_worker_thread_definitions_round_trip() {
        use vybe_ast::canon::ComponentCoreValType as T;
        let table = vec![0, 97, 115, 109, 1, 0, 0, 0,
            4, 4, 1, 0x70, 0, 1, 7, 5, 1, 1, b't', 1, 0];
        let component = ComponentBinary { definitions: vec![
            ComponentBinaryDef::CoreModule(table),
            ComponentBinaryDef::CoreInstance { module: 0, args: vec![] },
            ComponentBinaryDef::CoreAliasInstanceExport {
                instance: 0, name: "t".into(), kind: CoreExportKind::Table,
            },
            ComponentBinaryDef::CoreFuncType { params: vec![T::I32], results: vec![] },
            ComponentBinaryDef::CanonThreadIndex,
            ComponentBinaryDef::CanonThreadNewIndirect { type_index: 0, table_index: 0 },
        ] };
        let bytes = write_component_binary(&component).expect("valid worker thread definitions");
        assert_eq!(crate::read_component_binary(&bytes).unwrap(), component);
        for (type_index, table_index, expected) in [
            (1, 0, "core type index 1 is not declared"),
            (0, 1, "core table index 1 is not declared"),
        ] {
            let mut invalid = component.clone();
            *invalid.definitions.last_mut().unwrap() = ComponentBinaryDef::CanonThreadNewIndirect {
                type_index, table_index,
            };
            assert!(write_component_binary(&invalid).unwrap_err().contains(expected));
        }
    }

    #[test]
    fn core_reference_function_type_is_not_silently_rewritten() {
        let bytes = [0, 97, 115, 109, 13, 0, 1, 0,
            3, 5, 1, 0x60, 1, 0x6f, 0];
        assert!(crate::read_component_binary(&bytes).unwrap_err()
            .contains("unsupported component core value type"));
    }

    #[test]
    fn composite_value_types_round_trip_with_prior_type_references() {
        let p = |name: &str| ValSpec::Prim(name.into());
        let component = ComponentBinary {
            definitions: vec![
                ComponentBinaryDef::ValueType(p("u32")),
                ComponentBinaryDef::ValueType(ValSpec::List(Box::new(p("u32")))),
                ComponentBinaryDef::ValueType(ValSpec::Option(Box::new(ValSpec::Ref(1)))),
                ComponentBinaryDef::ValueType(ValSpec::Record(vec![
                    ("value".into(), ValSpec::Ref(2)),
                    ("name".into(), p("string")),
                ])),
                ComponentBinaryDef::ValueType(ValSpec::Variant(vec![
                    ("none".into(), None),
                    ("some".into(), Some(ValSpec::Ref(3))),
                ])),
                ComponentBinaryDef::ValueType(ValSpec::Tuple(vec![ValSpec::Ref(4), p("bool")])),
                ComponentBinaryDef::ValueType(ValSpec::Flags(vec!["a".into(), "b".into()])),
                ComponentBinaryDef::ValueType(ValSpec::Enum(vec!["x".into(), "y".into()])),
                ComponentBinaryDef::ValueType(ValSpec::Result(
                    Some(Box::new(ValSpec::Ref(5))),
                    Some(Box::new(p("string"))),
                )),
                ComponentBinaryDef::ValueType(ValSpec::Map(
                    Box::new(p("string")),
                    Box::new(ValSpec::Ref(8)),
                )),
                ComponentBinaryDef::ValueType(ValSpec::ListFixed(Box::new(p("u8")), 4)),
                ComponentBinaryDef::ValueType(ValSpec::Stream(Some(Box::new(ValSpec::Ref(9))))),
                ComponentBinaryDef::ValueType(ValSpec::Future(Some(Box::new(ValSpec::Ref(3))))),
                ComponentBinaryDef::ExportType {
                    name: "value".into(),
                    index: 3,
                },
            ],
        };
        let bytes = write_component_binary(&component).expect("composite component validates");
        let type_sections = wasmparser::Parser::new(0)
            .parse_all(&bytes)
            .filter(|payload| matches!(payload, Ok(wasmparser::Payload::ComponentTypeSection(_))))
            .count();
        assert_eq!(type_sections, 1);
        assert_eq!(crate::read_component_binary(&bytes).unwrap(), component);
    }

    #[test]
    fn composite_value_type_rejects_forward_reference() {
        let component = ComponentBinary {
            definitions: vec![ComponentBinaryDef::ValueType(ValSpec::Option(Box::new(
                ValSpec::Ref(1),
            )))],
        };
        assert!(
            write_component_binary(&component)
                .unwrap_err()
                .contains("not declared")
        );
    }

    #[test]
    fn value_type_reference_uses_signed_leb_past_index_63() {
        let mut definitions = vec![ComponentBinaryDef::ValueType(ValSpec::Prim("u32".into())); 65];
        definitions.push(ComponentBinaryDef::ValueType(ValSpec::Option(Box::new(
            ValSpec::Ref(64),
        ))));
        definitions.push(ComponentBinaryDef::ExportType {
            name: "high".into(),
            index: 65,
        });
        let component = ComponentBinary { definitions };
        let bytes = write_component_binary(&component).expect("signed type index validates");
        assert!(bytes.windows(3).any(|window| window == [0x6b, 0xc0, 0x00]));
        assert_eq!(crate::read_component_binary(&bytes).unwrap(), component);
    }

    #[test]
    fn sync_and_async_function_types_share_component_type_section() {
        let component = ComponentBinary {
            definitions: vec![
                ComponentBinaryDef::ValueType(ValSpec::Prim("u32".into())),
                ComponentBinaryDef::FuncType {
                    params: vec![("x".into(), ValSpec::Ref(0))],
                    result: Some(ValSpec::Prim("string".into())),
                    is_async: false,
                },
                ComponentBinaryDef::FuncType {
                    params: vec![("x".into(), ValSpec::Prim("u32".into()))],
                    result: Some(ValSpec::Prim("u32".into())),
                    is_async: true,
                },
                ComponentBinaryDef::FuncType {
                    params: Vec::new(),
                    result: None,
                    is_async: false,
                },
                ComponentBinaryDef::ExportType {
                    name: "plain".into(),
                    index: 1,
                },
            ],
        };
        let bytes = write_component_binary(&component).expect("function types validate");
        let type_sections = wasmparser::Parser::new(0)
            .parse_all(&bytes)
            .filter(|payload| matches!(payload, Ok(wasmparser::Payload::ComponentTypeSection(_))))
            .count();
        assert_eq!(type_sections, 1);
        assert_eq!(crate::read_component_binary(&bytes).unwrap(), component);
    }

    #[test]
    fn component_lengths_reject_u32_overflow() {
        assert_eq!(checked_len(u32::MAX as usize, "test").unwrap(), u32::MAX);
        if let Some(too_large) = (u32::MAX as usize).checked_add(1) {
            assert!(
                checked_len(too_large, "test")
                    .unwrap_err()
                    .contains("u32 limit")
            );
        }
    }

    #[test]
    fn large_custom_section_round_trips_without_changing_payload() {
        let component = ComponentBinary {
            definitions: vec![
                ComponentBinaryDef::Custom {
                    name: "metadata".into(),
                    data: (0..256 * 1024).map(|index| index as u8).collect(),
                },
                ComponentBinaryDef::CoreModule(vec![0, 0x61, 0x73, 0x6d, 1, 0, 0, 0]),
                ComponentBinaryDef::CoreInstance {
                    module: 0,
                    args: Vec::new(),
                },
                ComponentBinaryDef::Custom {
                    name: "tail".into(),
                    data: vec![0, 0xff, 1],
                },
            ],
        };
        let bytes = write_component_binary(&component).expect("component validates");
        assert_eq!(crate::read_component_binary(&bytes).unwrap(), component);
    }

    fn module_with_empty_functions(count: u32) -> Vec<u8> {
        let mut module = vec![0, 0x61, 0x73, 0x6d, 1, 0, 0, 0];
        write_section(&mut module, 1, &[1, 0x60, 0, 0]);
        let mut functions = Vec::new();
        write_leb128_u32(&mut functions, count);
        for _ in 0..count {
            functions.push(0);
        }
        write_section(&mut module, 3, &functions);
        let mut code = Vec::new();
        write_leb128_u32(&mut code, count);
        for _ in 0..count {
            code.extend_from_slice(&[2, 0, 0x0b]);
        }
        write_section(&mut module, 10, &code);
        module
    }

    #[test]
    fn empty_core_module_instantiates_as_real_component() {
        let component = ComponentBinary {
            definitions: vec![
                ComponentBinaryDef::CoreModule(vec![0, 0x61, 0x73, 0x6d, 1, 0, 0, 0]),
                ComponentBinaryDef::CoreInstance {
                    module: 0,
                    args: Vec::new(),
                },
            ],
        };
        let bytes = write_component_binary(&component).expect("component validates");
        assert_eq!(&bytes[..8], &[0, 0x61, 0x73, 0x6d, 0x0d, 0, 1, 0]);
        let mut modules = 0;
        let mut instances = 0;
        for payload in wasmparser::Parser::new(0).parse_all(&bytes) {
            match payload.expect("component payload") {
                wasmparser::Payload::ModuleSection { .. } => modules += 1,
                wasmparser::Payload::InstanceSection(reader) => {
                    for instance in reader {
                        assert!(matches!(
                            instance.expect("core instance"),
                            wasmparser::Instance::Instantiate {
                                module_index: 0,
                                ..
                            }
                        ));
                        instances += 1;
                    }
                }
                _ => {}
            }
        }
        assert_eq!((modules, instances), (1, 1));
        assert_eq!(crate::read_component_binary(&bytes).unwrap(), component);
    }

    #[test]
    fn core_instance_cannot_reference_future_module() {
        let component = ComponentBinary {
            definitions: vec![ComponentBinaryDef::CoreInstance {
                module: 0,
                args: Vec::new(),
            }],
        };
        assert!(
            write_component_binary(&component)
                .unwrap_err()
                .contains("not declared")
        );
    }

    #[test]
    fn segmented_payloads_match_contiguous_bytes_and_retain_final_validation() {
        for size in [65_535, 65_536] {
            let mut core = module_with_empty_functions(1);
            let mut custom = vec![0];
            custom.resize(size - 128, 42);
            write_section(&mut core, 0, &custom);
            let mut component = ComponentBinary { definitions: vec![
                ComponentBinaryDef::CoreModule(core),
                ComponentBinaryDef::Custom { name: "payload".into(), data: vec![7; size] },
                ComponentBinaryDef::FuncType { params: vec![], result: None, is_async: false },
                ComponentBinaryDef::ImportFunc { name: "f".into(), type_index: 0 },
                ComponentBinaryDef::ExportFunc { name: "f".into(), index: 0, type_index: None },
            ] };
            assert_eq!(has_bulk_payload(&component), size >= 65_536);
            for _ in 0..8 {
                component = ComponentBinary { definitions: vec![ComponentBinaryDef::NestedComponent(std::sync::Arc::new(component))] };
            }
            let reference = write_component_binary_impl(&component, true, true, true).unwrap();
            let bytes = write_component_binary(&component).unwrap();
            assert_eq!(bytes, reference);
            assert_eq!(crate::read_component_binary(&bytes).unwrap(), component);
        }
        let mut invalid = module_with_empty_functions(1);
        *invalid.last_mut().unwrap() = 0;
        let component = ComponentBinary { definitions: vec![
            ComponentBinaryDef::Custom { name: "payload".into(), data: vec![0; 65_536] },
            ComponentBinaryDef::NestedComponent(std::sync::Arc::new(ComponentBinary { definitions: vec![
                ComponentBinaryDef::CoreModule(invalid),
            ] })),
        ] };
        assert!(write_component_binary(&component).unwrap_err().contains("invalid component binary"));
    }

    #[test]
    fn invalid_embedded_core_module_is_rejected_by_component_validator() {
        let mut module = module_with_empty_functions(1);
        *module.last_mut().unwrap() = 0;
        let component = ComponentBinary {
            definitions: vec![ComponentBinaryDef::CoreModule(module)],
        };
        assert!(
            write_component_binary(&component)
                .unwrap_err()
                .contains("invalid component binary")
        );
    }

    #[test]
    fn invalid_deeply_nested_core_module_is_rejected() {
        let mut module = module_with_empty_functions(1);
        *module.last_mut().unwrap() = 0;
        let mut component = ComponentBinary {
            definitions: vec![ComponentBinaryDef::CoreModule(module)],
        };
        for _ in 0..8 {
            component = ComponentBinary {
                definitions: vec![ComponentBinaryDef::NestedComponent(std::sync::Arc::new(component))],
            };
        }
        assert!(write_component_binary(&component)
            .unwrap_err()
            .contains("invalid component binary"));
    }

    #[test]
    fn core_instance_with_import_uses_prior_instance() {
        let exporter = vec![
            0, 0x61, 0x73, 0x6d, 1, 0, 0, 0, 1, 4, 1, 0x60, 0, 0, 3, 2, 1, 0, 7, 5, 1, 1, b'f', 0,
            0, 10, 4, 1, 2, 0, 11,
        ];
        let importer = vec![
            0, 0x61, 0x73, 0x6d, 1, 0, 0, 0, 1, 4, 1, 0x60, 0, 0, 2, 7, 1, 1, b'm', 1, b'f', 0, 0,
            3, 2, 1, 0, 7, 5, 1, 1, b'g', 0, 1, 10, 6, 1, 4, 0, 0x10, 0, 11,
        ];
        let component = ComponentBinary {
            definitions: vec![
                ComponentBinaryDef::CoreModule(exporter),
                ComponentBinaryDef::CoreInstance {
                    module: 0,
                    args: Vec::new(),
                },
                ComponentBinaryDef::CoreModule(importer),
                ComponentBinaryDef::CoreInstance {
                    module: 1,
                    args: vec![("m".into(), 0)],
                },
            ],
        };
        let bytes = write_component_binary(&component).expect("linked component validates");
        assert_eq!(crate::read_component_binary(&bytes).unwrap(), component);
    }

    #[test]
    fn core_export_alias_can_populate_inline_instance() {
        let exporter = vec![
            0, 0x61, 0x73, 0x6d, 1, 0, 0, 0, 1, 4, 1, 0x60, 0, 0, 3, 2, 1, 0, 7, 5, 1, 1, b'f', 0,
            0, 10, 4, 1, 2, 0, 11,
        ];
        let component = ComponentBinary {
            definitions: vec![
                ComponentBinaryDef::CoreModule(exporter),
                ComponentBinaryDef::CoreInstance {
                    module: 0,
                    args: Vec::new(),
                },
                ComponentBinaryDef::CoreAliasInstanceExport {
                    instance: 0,
                    name: "f".into(),
                    kind: CoreExportKind::Func,
                },
                ComponentBinaryDef::CoreInstanceFromExports(vec![CoreInstanceExport {
                    name: "g".into(),
                    kind: CoreExportKind::Func,
                    index: 0,
                }]),
            ],
        };
        let bytes = write_component_binary(&component).expect("alias and inline instance validate");
        assert_eq!(crate::read_component_binary(&bytes).unwrap(), component);
    }

    #[test]
    fn canonical_lift_lower_round_trip() {
        let exporter = vec![
            0, 0x61, 0x73, 0x6d, 1, 0, 0, 0, 1, 4, 1, 0x60, 0, 0, 3, 2, 1, 0, 7, 5, 1, 1, b'f', 0,
            0, 10, 4, 1, 2, 0, 11,
        ];
        let component = ComponentBinary {
            definitions: vec![
                ComponentBinaryDef::CoreModule(exporter),
                ComponentBinaryDef::CoreInstance {
                    module: 0,
                    args: vec![],
                },
                ComponentBinaryDef::CoreAliasInstanceExport {
                    instance: 0,
                    name: "f".into(),
                    kind: CoreExportKind::Func,
                },
                ComponentBinaryDef::FuncType {
                    params: vec![],
                    result: None,
                    is_async: false,
                },
                ComponentBinaryDef::CanonLift {
                    core_func: 0,
                    type_index: 0,
                    options: vec![ComponentBinaryCanonOption::Utf8],
                },
                ComponentBinaryDef::CanonLower {
                    func: 0,
                    options: vec![],
                },
                ComponentBinaryDef::CoreInstanceFromExports(vec![CoreInstanceExport {
                    name: "lowered".into(),
                    kind: CoreExportKind::Func,
                    index: 1,
                }]),
            ],
        };
        let bytes = write_component_binary(&component).expect("canonical component validates");
        let canonical = wasmparser::Parser::new(0)
            .parse_all(&bytes)
            .find_map(|payload| match payload.ok()? {
                wasmparser::Payload::ComponentCanonicalSection(section) => {
                    let range = section.range();
                    let start = usize::try_from(range.start).expect("section offset within input");
                    let end = usize::try_from(range.end).expect("section offset within input");
                    Some(bytes[start..end].to_vec())
                }
                _ => None,
            })
            .expect("canonical section");
        assert_eq!(canonical, [2, 0, 0, 0, 1, 0, 0, 1, 0, 0, 0]);
        assert_eq!(crate::read_component_binary(&bytes).unwrap(), component);
    }

    #[test]
    fn canonical_task_return_matches_reference_binary() {
        let component = ComponentBinary {
            definitions: vec![ComponentBinaryDef::CanonTaskReturn {
                result: Some(ValSpec::Prim("u32".into())),
                options: vec![],
            }],
        };
        let bytes = write_component_binary(&component).expect("task.return validates");
        assert_eq!(bytes, [0, 0x61, 0x73, 0x6d, 0x0d, 0, 1, 0, 8, 5, 1, 9, 0, 0x79, 0]);
        assert_eq!(crate::read_component_binary(&bytes).unwrap(), component);

        let no_result = ComponentBinary {
            definitions: vec![ComponentBinaryDef::CanonTaskReturn {
                result: None,
                options: vec![],
            }],
        };
        let bytes = write_component_binary(&no_result).expect("result-less task.return validates");
        assert_eq!(crate::read_component_binary(&bytes).unwrap(), no_result);

        let indexed = ComponentBinary {
            definitions: vec![
                ComponentBinaryDef::ValueType(ValSpec::Prim("u32".into())),
                ComponentBinaryDef::CanonTaskReturn {
                    result: Some(ValSpec::Ref(0)),
                    options: vec![],
                },
            ],
        };
        let bytes = write_component_binary(&indexed).expect("indexed task.return validates");
        assert_eq!(crate::read_component_binary(&bytes).unwrap(), indexed);
    }

    #[test]
    fn imported_function_can_be_lowered_and_reexported() {
        let component = ComponentBinary {
            definitions: vec![
                ComponentBinaryDef::FuncType {
                    params: vec![],
                    result: None,
                    is_async: false,
                },
                ComponentBinaryDef::ImportFunc {
                    name: "f".into(),
                    type_index: 0,
                },
                ComponentBinaryDef::CanonLower {
                    func: 0,
                    options: vec![],
                },
                ComponentBinaryDef::ExportFunc {
                    name: "run".into(),
                    index: 0,
                    type_index: None,
                },
            ],
        };
        let bytes = write_component_binary(&component).expect("imported function validates");
        assert_eq!(
            bytes,
            vec![
                0, 0x61, 0x73, 0x6d, 0x0d, 0, 1, 0, 7, 5, 1, 0x40, 0, 1, 0, 10, 6, 1, 0, 1, b'f',
                1, 0, 8, 5, 1, 1, 0, 0, 0, 11, 9, 1, 0, 3, b'r', b'u', b'n', 1, 0, 0,
            ]
        );
        assert_eq!(crate::read_component_binary(&bytes).unwrap(), component);
    }

    #[test]
    fn adjacent_imports_share_sections_without_changing_function_indices() {
        let component = ComponentBinary {
            definitions: vec![
                ComponentBinaryDef::FuncType {
                    params: vec![],
                    result: None,
                    is_async: false,
                },
                ComponentBinaryDef::ImportFunc {
                    name: "a".into(),
                    type_index: 0,
                },
                ComponentBinaryDef::ImportFunc {
                    name: "b".into(),
                    type_index: 0,
                },
                ComponentBinaryDef::Custom {
                    name: "boundary".into(),
                    data: vec![],
                },
                ComponentBinaryDef::ImportFunc {
                    name: "c".into(),
                    type_index: 0,
                },
                ComponentBinaryDef::ExportFunc {
                    name: "third".into(),
                    index: 2,
                    type_index: None,
                },
            ],
        };
        let bytes = write_component_binary(&component).unwrap();
        let import_counts: Vec<_> = wasmparser::Parser::new(0)
            .parse_all(&bytes)
            .filter_map(|payload| match payload.unwrap() {
                wasmparser::Payload::ComponentImportSection(section) => Some(section.count()),
                _ => None,
            })
            .collect();
        assert_eq!(import_counts, [2, 1]);
        assert_eq!(crate::read_component_binary(&bytes).unwrap(), component);
    }

    #[test]
    fn nested_component_import_is_wired_through_instance_and_alias() {
        use vybe_ast::canon::{ComponentItemRef, ComponentItemKind};

        let function_type = || ComponentBinaryDef::FuncType {
            params: vec![],
            result: None,
            is_async: false,
        };
        let child = ComponentBinary {
            definitions: vec![
                function_type(),
                ComponentBinaryDef::ImportFunc {
                    name: "x".into(),
                    type_index: 0,
                },
                ComponentBinaryDef::ExportFunc {
                    name: "go".into(),
                    index: 0,
                    type_index: None,
                },
            ],
        };
        let parent = ComponentBinary {
            definitions: vec![
                function_type(),
                ComponentBinaryDef::ImportFunc {
                    name: "f".into(),
                    type_index: 0,
                },
                ComponentBinaryDef::NestedComponent(std::sync::Arc::new(child)),
                ComponentBinaryDef::ComponentInstance {
                    component: 0,
                    args: vec![ComponentItemRef {
                        name: "x".into(),
                        kind: ComponentItemKind::Func,
                        index: 0,
                    }],
                },
                ComponentBinaryDef::AliasInstanceExport {
                    instance: 0,
                    name: "go".into(),
                    kind: ComponentItemKind::Func,
                },
                ComponentBinaryDef::ExportFunc {
                    name: "run".into(),
                    index: 1,
                    type_index: None,
                },
            ],
        };
        let bytes = write_component_binary(&parent).expect("nested component validates");
        assert_eq!(crate::read_component_binary(&bytes).unwrap(), parent);
    }

    #[test]
    #[ignore]
    fn measure_large_core_module_component_write() {
        let component = ComponentBinary {
            definitions: vec![
                ComponentBinaryDef::CoreModule(module_with_empty_functions(20_000)),
                ComponentBinaryDef::CoreInstance {
                    module: 0,
                    args: Vec::new(),
                },
            ],
        };
        let mut timings = Vec::new();
        for _ in 0..7 {
            let start = std::time::Instant::now();
            let binary = write_component_binary(&component).unwrap();
            assert!(!binary.is_empty());
            timings.push(start.elapsed());
        }
        timings.sort();
        println!("large core-module component write median: {:?}", timings[3]);
    }

    #[test]
    #[ignore = "manual nested component writer benchmark"]
    fn benchmark_nested_component_write() {
        const DEPTH: usize = 8;
        const FUNCTIONS: u32 = 1024;
        const RUNS: usize = 30;
        let mut component = ComponentBinary {
            definitions: vec![ComponentBinaryDef::CoreModule(module_with_empty_functions(
                FUNCTIONS,
            ))],
        };
        for _ in 0..DEPTH {
            component = ComponentBinary {
                definitions: vec![ComponentBinaryDef::NestedComponent(std::sync::Arc::new(component))],
            };
        }
        let reference = write_component_binary_impl(&component, true, true, true).unwrap();
        let optimized = write_component_binary(&component).unwrap();
        assert_eq!(reference, optimized);
        let mut timings = [Vec::new(), Vec::new()];
        let mut measure = |mode: usize| {
            let start = std::time::Instant::now();
            let binary = if mode == 0 {
                write_component_binary_impl(&component, true, true, true).unwrap()
            } else {
                write_component_binary(&component).unwrap()
            };
            std::hint::black_box(binary);
            timings[mode].push(start.elapsed());
        };
        for run in 0..RUNS {
            measure(run % 2);
            measure(1 - run % 2);
        }
        for samples in &mut timings {
            samples.sort_unstable();
        }
        println!(
            "component writer: {DEPTH} nested components, {FUNCTIONS} funcs, \
             {RUNS} paired runs, recursive {:.3} ms, root-only {:.3} ms",
            timings[0][RUNS / 2].as_secs_f64() * 1000.0,
            timings[1][RUNS / 2].as_secs_f64() * 1000.0
        );
    }

    #[test]
    #[ignore = "manual component import-section benchmark"]
    fn benchmark_many_component_imports() {
        for mixed in [false, true] {
            measure_component_imports(mixed);
        }
    }

    fn measure_component_imports(mixed: bool) {
        const IMPORTS: usize = 8_000;
        const RUNS: usize = 7;
        let mut definitions = vec![ComponentBinaryDef::FuncType {
            params: vec![],
            result: None,
            is_async: false,
        }];
        if mixed {
            definitions.push(ComponentBinaryDef::InterfaceType {
                is_component: false, declarations: vec![],
            });
        }
        definitions.extend((0..IMPORTS).map(|index| {
            if mixed && index % 2 != 0 {
                ComponentBinaryDef::Import {
                    name: format!("i{index}"), ty: ComponentInterfaceRef::Instance(1),
                }
            } else {
                ComponentBinaryDef::ImportFunc { name: format!("f{index}"), type_index: 0 }
            }
        }));
        let component = ComponentBinary { definitions };
        let reference = write_component_binary_impl(&component, true, false, false).unwrap();
        let grouped = write_component_binary(&component).unwrap();
        assert_eq!(crate::read_component_binary(&reference).unwrap(), component);
        assert_eq!(crate::read_component_binary(&grouped).unwrap(), component);
        let mut times = [Vec::new(), Vec::new()];
        let mut measure = |mode: usize| {
            let start = std::time::Instant::now();
            let binary = if mode == 0 {
                write_component_binary_impl(&component, true, false, false).unwrap()
            } else {
                write_component_binary(&component).unwrap()
            };
            std::hint::black_box(binary);
            times[mode].push(start.elapsed());
        };
        for run in 0..RUNS {
            measure(run % 2);
            measure(1 - run % 2);
        }
        for samples in &mut times {
            samples.sort_unstable();
        }
        println!(
            "component imports: mixed={mixed}, {IMPORTS} imports, {} vs {} bytes, \
             separate {:.3} ms, grouped {:.3} ms",
            reference.len(),
            grouped.len(),
            times[0][RUNS / 2].as_secs_f64() * 1000.0,
            times[1][RUNS / 2].as_secs_f64() * 1000.0
        );
    }

    #[test]
    fn nested_interface_types_roundtrip_scoped_function_references() {
        use ComponentInterfaceDecl as D;
        use ComponentInterfaceRef as R;
        let component = ComponentBinary { definitions: vec![
            ComponentBinaryDef::FuncType {
                params: vec![], result: Some(ValSpec::Prim("u32".into())), is_async: false,
            },
            ComponentBinaryDef::InterfaceType { is_component: true, declarations: vec![
                D::AliasOuterType { count: 1, index: 0 },
                D::Import { name: "input".into(), ty: R::Func(0) },
                D::Type(Box::new(ComponentBinaryDef::InterfaceType {
                    is_component: false, declarations: vec![
                        D::AliasOuterType { count: 1, index: 0 },
                        D::Export { name: "value".into(), ty: R::Func(0) },
                    ],
                })),
                D::Export { name: "nested".into(), ty: R::Instance(1) },
            ] },
            ComponentBinaryDef::ExportType { name: "interface".into(), index: 1 },
        ] };
        let bytes = write_component_binary(&component).unwrap();
        assert_eq!(crate::read_component_binary(&bytes).unwrap(), component);
    }

    #[test]
    fn nested_component_clones_share_payload_and_preserve_copy_isolation() {
        use std::sync::Arc;
        let mut source = ComponentBinary { definitions: vec![
            ComponentBinaryDef::NestedComponent(Arc::new(ComponentBinary { definitions: vec![
                ComponentBinaryDef::CoreModule(module_with_empty_functions(3)),
            ] })),
        ] };
        let expected = write_component_binary(&source).unwrap();
        let copy = source.clone();
        let ComponentBinaryDef::NestedComponent(original_child) = &mut source.definitions[0]
            else { unreachable!() };
        let ComponentBinaryDef::NestedComponent(copied_child) = &copy.definitions[0]
            else { unreachable!() };
        assert!(Arc::ptr_eq(original_child, copied_child));
        Arc::make_mut(original_child).definitions.push(ComponentBinaryDef::Custom {
            name: "isolated".into(), data: vec![1, 2, 3],
        });
        assert!(!Arc::ptr_eq(original_child, copied_child));
        assert_eq!(write_component_binary(&copy).unwrap(), expected);
        let changed = write_component_binary(&source).unwrap();
        assert_ne!(changed, expected);
        assert_eq!(crate::read_component_binary(&changed).unwrap(), source);
        assert_eq!(crate::read_component_binary(&expected).unwrap(), copy);
    }

    #[test]
    fn interface_type_references_are_validated_by_sort() {
        use ComponentInterfaceDecl as D;
        use ComponentInterfaceRef as R;
        for is_component in [false, true] {
            let mut component = ComponentBinary { definitions: vec![
                ComponentBinaryDef::InterfaceType { is_component, declarations: vec![
                    D::Type(Box::new(ComponentBinaryDef::ValueType(ValSpec::Prim("u32".into())))),
                    D::Export { name: "value".into(), ty: R::Func(0) },
                ] },
            ] };
            assert!(write_component_binary(&component).unwrap_err().contains("not a function type"));
            let ComponentBinaryDef::InterfaceType { declarations, .. } = &mut component.definitions[0]
                else { unreachable!() };
            declarations[0] = D::Type(Box::new(ComponentBinaryDef::FuncType {
                params: vec![], result: None, is_async: false,
            }));
            let bytes = write_component_binary(&component).unwrap();
            assert_eq!(crate::read_component_binary(&bytes).unwrap(), component);
        }
    }

    #[test]
    fn mixed_component_imports_preserve_sort_and_index_spaces() {
        use ComponentInterfaceRef as R;
        let component = ComponentBinary { definitions: vec![
            ComponentBinaryDef::InterfaceType { is_component: false, declarations: vec![] },
            ComponentBinaryDef::InterfaceType { is_component: true, declarations: vec![] },
            ComponentBinaryDef::FuncType { params: vec![], result: None, is_async: false },
            ComponentBinaryDef::Import { name: "instance".into(), ty: R::Instance(0) },
            ComponentBinaryDef::Import { name: "component".into(), ty: R::Component(1) },
            ComponentBinaryDef::Import { name: "resource".into(), ty: R::TypeSubResource },
            ComponentBinaryDef::ImportFunc { name: "function".into(), type_index: 2 },
            ComponentBinaryDef::ExportInstance { name: "instance".into(), index: 0 },
            ComponentBinaryDef::ExportInstance { name: "instance-again".into(), index: 1 },
            ComponentBinaryDef::ExportType { name: "resource".into(), index: 3 },
            ComponentBinaryDef::ExportFunc { name: "function".into(), index: 0, type_index: None },
        ] };
        let bytes = write_component_binary(&component).unwrap();
        assert_eq!(component, crate::read_component_binary(&bytes).unwrap());
        let counts: Vec<_> = wasmparser::Parser::new(0).parse_all(&bytes)
            .filter_map(|payload| match payload.unwrap() {
                wasmparser::Payload::ComponentImportSection(section) => Some(section.count()),
                _ => None,
            }).collect();
        assert_eq!(counts, vec![4]);
        let ungrouped = write_component_binary_impl(&component, true, false, false).unwrap();
        assert_eq!(component, crate::read_component_binary(&ungrouped).unwrap());
    }

    #[test]
    fn inline_instances_retain_all_component_item_sorts() {
        use vybe_ast::canon::{ComponentItemKind as K, ComponentItemRef as I};
        let item = |name: &str, kind, index| I { name: name.into(), kind, index };
        let mut component = ComponentBinary { definitions: vec![
            ComponentBinaryDef::FuncType { params: vec![], result: None, is_async: false },
            ComponentBinaryDef::InterfaceType { is_component: false, declarations: vec![] },
            ComponentBinaryDef::InterfaceType { is_component: true, declarations: vec![] },
            ComponentBinaryDef::ImportFunc { name: "function".into(), type_index: 0 },
            ComponentBinaryDef::Import { name: "instance".into(), ty: ComponentInterfaceRef::Instance(1) },
            ComponentBinaryDef::Import { name: "component".into(), ty: ComponentInterfaceRef::Component(2) },
            ComponentBinaryDef::Import { name: "resource".into(), ty: ComponentInterfaceRef::TypeSubResource },
            ComponentBinaryDef::Import { name: "value".into(), ty: ComponentInterfaceRef::Value(ValSpec::Prim("u32".into())) },
            ComponentBinaryDef::CoreModule(vec![0, 0x61, 0x73, 0x6d, 1, 0, 0, 0]),
            ComponentBinaryDef::InstanceFromExports(vec![
                item("module", K::Module, 0), item("function", K::Func, 0),
                item("instance", K::Instance, 0), item("component", K::Component, 0),
                item("resource", K::Type, 3), item("value", K::Value, 0),
            ]),
            ComponentBinaryDef::ExportInstance { name: "assembled".into(), index: 1 },
        ] };
        let bytes = write_component_binary(&component).unwrap();
        assert_eq!(component, crate::read_component_binary(&bytes).unwrap());
        let ComponentBinaryDef::InstanceFromExports(exports) = &mut component.definitions[9]
            else { unreachable!() };
        exports[1].index = 1;
        assert!(write_component_binary(&component).unwrap_err().contains("Func index 1 is not declared"));
    }

    #[test]
    fn empty_inline_instance_matches_standard_wire_bytes() {
        let component = ComponentBinary { definitions: vec![
            ComponentBinaryDef::InstanceFromExports(vec![]),
        ] };
        let expected = vec![
            0, 0x61, 0x73, 0x6d, 0x0d, 0, 1, 0,
            5, 3, 1, 1, 0,
        ];
        assert_eq!(write_component_binary(&component).unwrap(), expected);
        assert_eq!(crate::read_component_binary(&expected).unwrap(), component);
    }

    #[test]
    fn grouped_instances_preserve_prior_references_and_section_boundaries() {
        use vybe_ast::canon::{ComponentItemKind as K, ComponentItemRef as I};
        let export = |index| I { name: "prior".into(), kind: K::Instance, index };
        let component = ComponentBinary { definitions: vec![
            ComponentBinaryDef::InstanceFromExports(vec![]),
            ComponentBinaryDef::InstanceFromExports(vec![export(0)]),
            ComponentBinaryDef::InstanceFromExports(vec![]),
            ComponentBinaryDef::Custom { name: "boundary".into(), data: vec![1] },
            ComponentBinaryDef::InstanceFromExports(vec![export(1)]),
            ComponentBinaryDef::ExportInstance { name: "result".into(), index: 3 },
        ] };
        let bytes = write_component_binary(&component).unwrap();
        assert_eq!(component, crate::read_component_binary(&bytes).unwrap());
        let counts: Vec<_> = wasmparser::Parser::new(0).parse_all(&bytes)
            .filter_map(|payload| match payload.unwrap() {
                wasmparser::Payload::ComponentInstanceSection(section) => Some(section.count()),
                _ => None,
            }).collect();
        assert_eq!(counts, vec![3, 1]);
        let reference = write_component_binary_impl(&component, true, false, false).unwrap();
        assert_eq!(component, crate::read_component_binary(&reference).unwrap());
    }

    #[test]
    #[ignore = "manual grouped component instance benchmark"]
    fn benchmark_many_component_instances() {
        const INSTANCES: usize = 4_000;
        const RUNS: usize = 7;
        let mut definitions = vec![ComponentBinaryDef::NestedComponent(std::sync::Arc::new(ComponentBinary::default()))];
        definitions.extend((0..INSTANCES).map(|index| if index % 2 == 0 {
            ComponentBinaryDef::InstanceFromExports(vec![])
        } else {
            ComponentBinaryDef::ComponentInstance { component: 0, args: vec![] }
        }));
        let component = ComponentBinary { definitions };
        let reference = write_component_binary_impl(&component, true, false, false).unwrap();
        let grouped = write_component_binary(&component).unwrap();
        assert_eq!(component, crate::read_component_binary(&reference).unwrap());
        assert_eq!(component, crate::read_component_binary(&grouped).unwrap());
        let mut times = [Vec::new(), Vec::new()];
        for run in 0..RUNS {
            for mode in [run % 2, 1 - run % 2] {
                let start = std::time::Instant::now();
                let binary = if mode == 0 {
                    write_component_binary_impl(&component, true, false, false).unwrap()
                } else { write_component_binary(&component).unwrap() };
                std::hint::black_box(binary);
                times[mode].push(start.elapsed());
            }
        }
        for samples in &mut times { samples.sort_unstable(); }
        println!("component instances: {INSTANCES}, {} vs {} bytes, separate {:.3} ms, grouped {:.3} ms",
            reference.len(), grouped.len(), times[0][RUNS / 2].as_secs_f64() * 1000.0,
            times[1][RUNS / 2].as_secs_f64() * 1000.0);
    }

    #[test]
    fn instance_arguments_validate_sort_bounds_and_repeated_instantiation() {
        use vybe_ast::canon::{ComponentItemKind as K, ComponentItemRef as I};
        let child = ComponentBinary { definitions: vec![
            ComponentBinaryDef::InterfaceType { is_component: false, declarations: vec![] },
            ComponentBinaryDef::Import { name: "provider".into(), ty: ComponentInterfaceRef::Instance(0) },
        ] };
        let args = vec![I { name: "provider".into(), kind: K::Instance, index: 0 }];
        let mut component = ComponentBinary { definitions: vec![
            ComponentBinaryDef::FuncType { params: vec![], result: None, is_async: false },
            ComponentBinaryDef::ImportFunc { name: "function".into(), type_index: 0 },
            ComponentBinaryDef::InstanceFromExports(vec![]),
            ComponentBinaryDef::NestedComponent(std::sync::Arc::new(child)),
            ComponentBinaryDef::ComponentInstance { component: 0, args: args.clone() },
            ComponentBinaryDef::ComponentInstance { component: 0, args },
        ] };
        let bytes = write_component_binary(&component).unwrap();
        assert_eq!(component, crate::read_component_binary(&bytes).unwrap());
        let ComponentBinaryDef::ComponentInstance { args, .. } = &mut component.definitions[4]
            else { unreachable!() };
        args[0].kind = K::Func;
        assert!(write_component_binary(&component).is_err());
        let ComponentBinaryDef::ComponentInstance { args, .. } = &mut component.definitions[4]
            else { unreachable!() };
        args[0].kind = K::Instance;
        args[0].index = 1;
        assert!(write_component_binary(&component).unwrap_err().contains("Instance index 1 is not declared"));
        let ComponentBinaryDef::ComponentInstance { args, .. } = &mut component.definitions[4]
            else { unreachable!() };
        args[0].index = 0;
        args.push(args[0].clone());
        assert!(write_component_binary(&component).is_err());
    }

    #[test]
    fn component_imports_validate_type_sort_and_resource_bounds() {
        use ComponentInterfaceRef as R;
        for ty in [R::Instance(0), R::Component(0)] {
            let component = ComponentBinary { definitions: vec![
                ComponentBinaryDef::FuncType { params: vec![], result: None, is_async: false },
                ComponentBinaryDef::Import { name: "wrong".into(), ty },
            ] };
            assert!(write_component_binary(&component).is_err());
        }
        let component = ComponentBinary { definitions: vec![
            ComponentBinaryDef::Import { name: "resource".into(), ty: R::TypeSubResource },
            ComponentBinaryDef::ValueType(ValSpec::Own(0)),
            ComponentBinaryDef::FuncType { params: vec![("value".into(), ValSpec::Ref(1))], result: None, is_async: false },
            ComponentBinaryDef::ImportFunc { name: "consume".into(), type_index: 2 },
        ] };
        let bytes = write_component_binary(&component).unwrap();
        assert_eq!(component, crate::read_component_binary(&bytes).unwrap());
    }

    #[test]
    fn interface_resource_bounds_advance_local_type_space() {
        use ComponentInterfaceDecl as D;
        use ComponentInterfaceRef as R;
        let component = ComponentBinary { definitions: vec![
            ComponentBinaryDef::InterfaceType { is_component: true, declarations: vec![
                D::Import { name: "resource".into(), ty: R::TypeSubResource },
                D::Type(Box::new(ComponentBinaryDef::ValueType(ValSpec::Own(0)))),
                D::Type(Box::new(ComponentBinaryDef::FuncType {
                    params: vec![("value".into(), ValSpec::Ref(1))], result: None, is_async: false,
                })),
                D::Export { name: "consume".into(), ty: R::Func(2) },
            ] },
        ] };
        let bytes = write_component_binary(&component).unwrap();
        assert_eq!(crate::read_component_binary(&bytes).unwrap(), component);
    }
}
