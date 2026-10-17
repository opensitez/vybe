//! Checked component-layer binary decoding into the shared typed IR.
use super::{WasmError, section_types};
use crate::binary_cursor::Cursor;
use crate::binary_sections::{BinaryKind, Sections};
use vybe_ast::canon::{
    ComponentBinary, ComponentBinaryAsyncKind as AK, ComponentBinaryAsyncOperation as AO,
    ComponentBinaryCanonOption as O, ComponentBinaryDef as D, ComponentCoreValType as CV,
    ComponentInterfaceDecl as I, ComponentInterfaceRef as T, ComponentItemKind as K,
    ComponentItemRef, CoreExportKind, CoreInstanceExport, ResourceRep, ValSpec,
};
type R<T> = Result<T, WasmError>;

fn name(cursor: &mut Cursor<'_>) -> R<String> {
    Ok(cursor.name()?.to_owned())
}
fn zero(cursor: &mut Cursor<'_>) -> R<()> {
    if cursor.byte()? != 0 {
        return Err("expected reserved zero byte".into());
    }
    Ok(())
}
fn boolean(cursor: &mut Cursor<'_>) -> R<bool> {
    match cursor.byte()? {
        0 => Ok(false),
        1 => Ok(true),
        _ => Err("invalid boolean byte".into()),
    }
}
fn external_name(cursor: &mut Cursor<'_>) -> R<String> {
    match cursor.byte()? {
        0 | 1 => name(cursor),
        2 => {
            let name = name(cursor)?;
            if cursor.u32()? != 0 {
                return Err(
                    "component external name metadata is not represented in binary IR".into(),
                );
            }
            Ok(name)
        }
        _ => Err("component external name metadata is not represented in binary IR".into()),
    }
}
fn valtype(cursor: &mut Cursor<'_>) -> R<ValSpec> {
    if let Some(name) = cursor
        .peek()
        .and_then(crate::component_format::primitive_name)
    {
        cursor.byte()?;
        return Ok(ValSpec::Prim(name.into()));
    }
    let index = u32::try_from(cursor.s33()?)
        .map_err(|_| WasmError::from("invalid component value type"))?;
    Ok(ValSpec::Ref(index))
}
fn optional_valtype(cursor: &mut Cursor<'_>) -> R<Option<Box<ValSpec>>> {
    if boolean(cursor)? {
        Ok(Some(Box::new(valtype(cursor)?)))
    } else {
        Ok(None)
    }
}
fn result_type(cursor: &mut Cursor<'_>) -> R<Option<ValSpec>> {
    match cursor.byte()? {
        0 => Ok(Some(valtype(cursor)?)),
        1 => {
            if cursor.u32()? != 0 {
                return Err("named component results are not supported".into());
            }
            Ok(None)
        }
        _ => Err("invalid component result encoding".into()),
    }
}
fn value_definition(cursor: &mut Cursor<'_>) -> R<ValSpec> {
    let tag = cursor.byte()?;
    if let Some(name) = crate::component_format::primitive_name(tag) {
        return Ok(ValSpec::Prim(name.into()));
    }
    Ok(match tag {
        0x72 => {
            let count = section_types::count(cursor, 1)?;
            let mut fields = Vec::with_capacity(count);
            for _ in 0..count {
                fields.push((name(cursor)?, valtype(cursor)?));
            }
            ValSpec::Record(fields)
        }
        0x71 => {
            let count = section_types::count(cursor, 1)?;
            let mut cases = Vec::with_capacity(count);
            for _ in 0..count {
                let name = name(cursor)?;
                let ty = optional_valtype(cursor)?.map(|v| *v);
                zero(cursor)?;
                cases.push((name, ty));
            }
            ValSpec::Variant(cases)
        }
        0x70 => ValSpec::List(Box::new(valtype(cursor)?)),
        0x67 => {
            let ty = Box::new(valtype(cursor)?);
            ValSpec::ListFixed(ty, cursor.u32()?)
        }
        0x6f => {
            let count = section_types::count(cursor, 1)?;
            let mut fields = Vec::with_capacity(count);
            for _ in 0..count {
                fields.push(valtype(cursor)?);
            }
            ValSpec::Tuple(fields)
        }
        0x6e | 0x6d => {
            let count = section_types::count(cursor, 1)?;
            let mut names = Vec::with_capacity(count);
            for _ in 0..count {
                names.push(name(cursor)?);
            }
            if tag == 0x6e {
                ValSpec::Flags(names)
            } else {
                ValSpec::Enum(names)
            }
        }
        0x6b => ValSpec::Option(Box::new(valtype(cursor)?)),
        0x6a => ValSpec::Result(optional_valtype(cursor)?, optional_valtype(cursor)?),
        0x69 => ValSpec::Own(cursor.u32()?),
        0x68 => ValSpec::Borrow(cursor.u32()?),
        0x66 => ValSpec::Stream(optional_valtype(cursor)?),
        0x65 => ValSpec::Future(optional_valtype(cursor)?),
        0x63 => ValSpec::Map(Box::new(valtype(cursor)?), Box::new(valtype(cursor)?)),
        _ => return Err(format!("unsupported component value type: {tag:#x}").into()),
    })
}
fn item_kind(cursor: &mut Cursor<'_>) -> R<K> {
    Ok(match cursor.byte()? {
        0 => {
            if cursor.byte()? != 0x11 {
                return Err("invalid component core sort".into());
            }
            K::Module
        }
        1 => K::Func,
        2 => K::Value,
        3 => K::Type,
        4 => K::Component,
        5 => K::Instance,
        _ => return Err("invalid component item sort".into()),
    })
}
fn core_kind(byte: u8) -> R<CoreExportKind> {
    Ok(match byte {
        0 => CoreExportKind::Func,
        1 => CoreExportKind::Table,
        2 => CoreExportKind::Memory,
        3 => CoreExportKind::Global,
        4 => CoreExportKind::Tag,
        _ => return Err("invalid core export kind".into()),
    })
}
fn interface_ref(cursor: &mut Cursor<'_>) -> R<T> {
    Ok(match item_kind(cursor)? {
        K::Module => T::Module(cursor.u32()?),
        K::Func => T::Func(cursor.u32()?),
        K::Value => T::Value(valtype(cursor)?),
        K::Type => match cursor.byte()? {
            0 => T::TypeEq(cursor.u32()?),
            1 => T::TypeSubResource,
            _ => return Err("invalid component type bound".into()),
        },
        K::Component => T::Component(cursor.u32()?),
        K::Instance => T::Instance(cursor.u32()?),
    })
}
enum Alias {
    Instance {
        instance: u32,
        name: String,
        kind: K,
    },
    Core {
        instance: u32,
        name: String,
        kind: CoreExportKind,
    },
    OuterType {
        count: u32,
        index: u32,
    },
}
fn alias(cursor: &mut Cursor<'_>) -> R<Alias> {
    let first = cursor.byte()?;
    let (component_sort, core_sort) = if first == 0 {
        let second = cursor.byte()?;
        if second == 0x11 {
            (Some(K::Module), None)
        } else {
            (None, Some(core_kind(second)?))
        }
    } else {
        let kind = match first {
            1 => K::Func,
            2 => K::Value,
            3 => K::Type,
            4 => K::Component,
            5 => K::Instance,
            _ => return Err("invalid alias sort".into()),
        };
        (Some(kind), None)
    };
    match cursor.byte()? {
        0 => Ok(Alias::Instance {
            instance: cursor.u32()?,
            name: name(cursor)?,
            kind: component_sort.ok_or("component alias has core sort")?,
        }),
        1 => Ok(Alias::Core {
            instance: cursor.u32()?,
            name: name(cursor)?,
            kind: core_sort.ok_or("core alias has component sort")?,
        }),
        2 if component_sort == Some(K::Type) => Ok(Alias::OuterType {
            count: cursor.u32()?,
            index: cursor.u32()?,
        }),
        _ => Err("unsupported component alias".into()),
    }
}
fn type_definition(cursor: &mut Cursor<'_>) -> R<D> {
    match cursor.peek() {
        Some(0x40 | 0x43) => {
            let is_async = cursor.byte()? == 0x43;
            let count = section_types::count(cursor, 1)?;
            let mut params = Vec::with_capacity(count);
            for _ in 0..count {
                params.push((name(cursor)?, valtype(cursor)?));
            }
            Ok(D::FuncType {
                params,
                result: result_type(cursor)?,
                is_async,
            })
        }
        Some(0x3f) => {
            cursor.byte()?;
            let rep = match cursor.byte()? {
                0x7f => ResourceRep::I32,
                0x7e => ResourceRep::I64,
                _ => return Err("unsupported resource representation".into()),
            };
            let dtor = if boolean(cursor)? {
                Some(cursor.u32()?)
            } else {
                None
            };
            Ok(D::ResourceType { rep, dtor })
        }
        Some(0x41 | 0x42) => {
            let is_component = cursor.byte()? == 0x41;
            let count = section_types::count(cursor, 1)?;
            let mut declarations = Vec::with_capacity(count);
            for _ in 0..count {
                declarations.push(match cursor.byte()? {
                    0 => return Err("interface core type is not represented in binary IR".into()),
                    1 => I::Type(Box::new(type_definition(cursor)?)),
                    2 => match alias(cursor)? {
                        Alias::OuterType { count, index } => I::AliasOuterType { count, index },
                        Alias::Instance {
                            instance,
                            name,
                            kind,
                        } => I::AliasInstanceExport {
                            instance,
                            name,
                            kind,
                        },
                        _ => return Err("unsupported interface type alias".into()),
                    },
                    3 if is_component => I::Import {
                        name: external_name(cursor)?,
                        ty: interface_ref(cursor)?,
                    },
                    4 => I::Export {
                        name: external_name(cursor)?,
                        ty: interface_ref(cursor)?,
                    },
                    _ => return Err("invalid interface declaration".into()),
                });
            }
            Ok(D::InterfaceType {
                is_component,
                declarations,
            })
        }
        _ => Ok(D::ValueType(value_definition(cursor)?)),
    }
}
fn core_valtype(cursor: &mut Cursor<'_>) -> R<CV> {
    Ok(match cursor.byte()? {
        0x7f => CV::I32,
        0x7e => CV::I64,
        0x7d => CV::F32,
        0x7c => CV::F64,
        0x7b => CV::V128,
        _ => return Err("unsupported component core value type".into()),
    })
}
fn core_function_type(cursor: &mut Cursor<'_>) -> R<D> {
    if cursor.peek() == Some(0x4f) {
        cursor.byte()?;
        if cursor.u32()? != 0 {
            return Err("unsupported component core subtype metadata".into());
        }
    }
    if cursor.byte()? != 0x60 {
        return Err("unsupported component core non-function or recursive type".into());
    }
    let count = section_types::count(cursor, 1)?;
    let mut params = Vec::with_capacity(count);
    for _ in 0..count {
        params.push(core_valtype(cursor)?);
    }
    let count = section_types::count(cursor, 1)?;
    let mut results = Vec::with_capacity(count);
    for _ in 0..count {
        results.push(core_valtype(cursor)?);
    }
    Ok(D::CoreFuncType { params, results })
}
fn options(cursor: &mut Cursor<'_>) -> R<Vec<O>> {
    let count = section_types::count(cursor, 1)?;
    let mut result = Vec::with_capacity(count);
    for _ in 0..count {
        result.push(match cursor.byte()? {
            0 => O::Utf8,
            1 => O::Utf16,
            2 => O::CompactUtf16,
            3 => O::Memory(cursor.u32()?),
            4 => O::Realloc(cursor.u32()?),
            5 => O::PostReturn(cursor.u32()?),
            6 => O::Async,
            7 => O::Callback(cursor.u32()?),
            8 => O::CoreType(cursor.u32()?),
            9 => O::Gc,
            _ => return Err("invalid canonical option".into()),
        });
    }
    Ok(result)
}
fn canonical(cursor: &mut Cursor<'_>) -> R<D> {
    let tag = cursor.byte()?;
    Ok(match tag {
        0 => {
            zero(cursor)?;
            D::CanonLift {
                core_func: cursor.u32()?,
                options: options(cursor)?,
                type_index: cursor.u32()?,
            }
        }
        1 => {
            zero(cursor)?;
            D::CanonLower {
                func: cursor.u32()?,
                options: options(cursor)?,
            }
        }
        2 => D::CanonResourceNew {
            type_index: cursor.u32()?,
        },
        3 => D::CanonResourceDrop {
            type_index: cursor.u32()?,
        },
        4 => D::CanonResourceRep {
            type_index: cursor.u32()?,
        },
        5 => D::CanonTaskCancel,
        6 => D::CanonSubtaskCancel {
            is_async: boolean(cursor)?,
        },
        9 => D::CanonTaskReturn {
            result: result_type(cursor)?,
            options: options(cursor)?,
        },
        0x0a | 0x0b => {
            let ty = core_valtype(cursor)?;
            if !matches!(ty, CV::I32 | CV::I64) {
                return Err("canonical context type must be i32 or i64".into());
            }
            D::CanonContext {
                set: tag == 0x0b,
                ty,
                slot: cursor.u32()?,
            }
        }
        0x0d => D::CanonSubtaskDrop,
        0x0e..=0x1b | 0x2e | 0x2f => {
            let (kind, offset) = match tag {
                0x0e..=0x14 => (AK::Stream, tag - 0x0e),
                0x15..=0x1b => (AK::Future, tag - 0x15),
                0x2e => (AK::Stream, 7),
                _ => (AK::Future, 7),
            };
            let type_index = cursor.u32()?;
            let operation = match offset {
                0 => AO::New,
                1 => AO::Read,
                2 => AO::Write,
                3 => AO::CancelRead {
                    is_async: boolean(cursor)?,
                },
                4 => AO::CancelWrite {
                    is_async: boolean(cursor)?,
                },
                5 => AO::DropReadable,
                6 => AO::DropWritable,
                _ => AO::Forward,
            };
            let options = if matches!(operation, AO::Read | AO::Write) {
                options(cursor)?
            } else {
                Vec::new()
            };
            D::CanonAsync {
                kind,
                operation,
                type_index,
                options,
            }
        }
        0x1c => D::CanonErrorContextNew {
            options: options(cursor)?,
        },
        0x1d => D::CanonErrorContextDebugMessage {
            options: options(cursor)?,
        },
        0x1e => D::CanonErrorContextDrop,
        0x1f => D::CanonWaitableSetNew,
        0x20 | 0x21 => {
            zero(cursor)?;
            let memory = cursor.u32()?;
            if tag == 0x20 {
                D::CanonWaitableSetWait { memory }
            } else {
                D::CanonWaitableSetPoll { memory }
            }
        }
        0x22 => D::CanonWaitableSetDrop,
        0x23 => D::CanonWaitableJoin,
        0x24 => D::CanonBackpressureInc,
        0x25 => D::CanonBackpressureDec,
        0x26 => D::CanonThreadIndex,
        0x27 => D::CanonThreadNewIndirect {
            type_index: cursor.u32()?,
            table_index: cursor.u32()?,
        },
        0x28 => D::CanonThreadResumeLater,
        0x0c | 0x29..=0x2d => {
            zero(cursor)?;
            match tag {
                0x0c => D::CanonThreadYield,
                0x29 => D::CanonThreadSuspend,
                0x2a => D::CanonThreadSuspendThenResume,
                0x2b => D::CanonThreadYieldThenResume,
                0x2c => D::CanonThreadSuspendThenPromote,
                _ => D::CanonThreadYieldThenPromote,
            }
        }
        _ => return Err(format!("unsupported canonical definition: {tag:#x}").into()),
    })
}
fn item(cursor: &mut Cursor<'_>, name: String) -> R<ComponentItemRef> {
    Ok(ComponentItemRef {
        name,
        kind: item_kind(cursor)?,
        index: cursor.u32()?,
    })
}
fn core_instance(cursor: &mut Cursor<'_>) -> R<D> {
    match cursor.byte()? {
        0 => {
            let module = cursor.u32()?;
            let count = section_types::count(cursor, 1)?;
            let mut args = Vec::with_capacity(count);
            for _ in 0..count {
                let name = name(cursor)?;
                if cursor.byte()? != 0x12 {
                    return Err("invalid core instantiation argument sort".into());
                }
                args.push((name, cursor.u32()?));
            }
            Ok(D::CoreInstance { module, args })
        }
        1 => {
            let count = section_types::count(cursor, 1)?;
            let mut exports = Vec::with_capacity(count);
            for _ in 0..count {
                exports.push(CoreInstanceExport {
                    name: name(cursor)?,
                    kind: core_kind(cursor.byte()?)?,
                    index: cursor.u32()?,
                });
            }
            Ok(D::CoreInstanceFromExports(exports))
        }
        _ => Err("invalid core instance".into()),
    }
}
fn component_instance(cursor: &mut Cursor<'_>) -> R<D> {
    match cursor.byte()? {
        0 => {
            let component = cursor.u32()?;
            let count = section_types::count(cursor, 1)?;
            let mut args = Vec::with_capacity(count);
            for _ in 0..count {
                let name = name(cursor)?;
                args.push(item(cursor, name)?);
            }
            Ok(D::ComponentInstance { component, args })
        }
        1 => {
            let count = section_types::count(cursor, 1)?;
            let mut exports = Vec::with_capacity(count);
            for _ in 0..count {
                let name = external_name(cursor)?;
                exports.push(item(cursor, name)?);
            }
            Ok(D::InstanceFromExports(exports))
        }
        _ => Err("invalid component instance".into()),
    }
}
fn export(cursor: &mut Cursor<'_>) -> R<D> {
    let name = external_name(cursor)?;
    let kind = item_kind(cursor)?;
    let index = cursor.u32()?;
    let ty = if boolean(cursor)? {
        Some(interface_ref(cursor)?)
    } else {
        None
    };
    if kind != K::Func && ty.is_some() {
        return Err("non-function export ascription is not represented in binary IR".into());
    }
    Ok(match kind {
        K::Module => D::ExportCoreModule { name, index },
        K::Instance => D::ExportInstance { name, index },
        K::Type => D::ExportType { name, index },
        K::Func => D::ExportFunc {
            name,
            index,
            type_index: match ty {
                None => None,
                Some(T::Func(index)) => Some(index),
                _ => return Err("function export has a non-function ascription".into()),
            },
        },
        _ => return Err("unsupported component export".into()),
    })
}
fn component_sections(bytes: &[u8]) -> R<Sections<'_>> {
    let sections = Sections::new(bytes)?;
    if sections.kind() != BinaryKind::Component {
        return Err("binary is not a supported component".into());
    }
    Ok(sections)
}
fn decode(bytes: &[u8]) -> R<ComponentBinary> {
    // Explicit frames avoid native stack growth for deeply nested components.
    let mut frames = vec![(component_sections(bytes)?, ComponentBinary::default())];
    loop {
        let next = frames.last_mut().unwrap().0.next();
        let Some(section) = next else {
            let (_, finished) = frames.pop().unwrap();
            if let Some((_, parent)) = frames.last_mut() {
                parent
                    .definitions
                    .push(D::NestedComponent(std::sync::Arc::new(finished)));
                continue;
            }
            return Ok(finished);
        };
        let section = section?;
        let id = section.id;
        let payload = section.payload;
        let component = &mut frames.last_mut().unwrap().1;
        match id {
            0 => {
                let mut custom = Cursor::new(payload);
                let name = name(&mut custom)?;
                let data = custom.bytes(custom.remaining())?.to_vec();
                component.definitions.push(D::Custom { name, data });
            }
            1 => component.definitions.push(D::CoreModule(payload.to_vec())),
            4 => {
                frames.push((component_sections(payload)?, ComponentBinary::default()));
            }
            2 | 3 | 5..=8 | 10 | 11 => {
                let mut section = Cursor::new(payload);
                let count = section_types::count(&mut section, 1)?;
                for _ in 0..count {
                    let definition = match id {
                        2 => core_instance(&mut section)?,
                        3 => core_function_type(&mut section)?,
                        5 => component_instance(&mut section)?,
                        6 => match alias(&mut section)? {
                            Alias::Instance {
                                instance,
                                name,
                                kind,
                            } => D::AliasInstanceExport {
                                instance,
                                name,
                                kind,
                            },
                            Alias::Core {
                                instance,
                                name,
                                kind,
                            } => D::CoreAliasInstanceExport {
                                instance,
                                name,
                                kind,
                            },
                            _ => return Err("unsupported component outer alias".into()),
                        },
                        7 => type_definition(&mut section)?,
                        8 => canonical(&mut section)?,
                        10 => {
                            let name = external_name(&mut section)?;
                            let ty = interface_ref(&mut section)?;
                            if let T::Func(type_index) = ty {
                                D::ImportFunc { name, type_index }
                            } else {
                                D::Import { name, ty }
                            }
                        }
                        11 => export(&mut section)?,
                        _ => unreachable!(),
                    };
                    component.definitions.push(definition);
                }
                section_types::finish(&section)?;
            }
            _ => return Err(format!("unsupported component section: {id}").into()),
        }
    }
}

pub fn read_component_binary(bytes: &[u8]) -> Result<ComponentBinary, String> {
    super::native_validation::validate_component(bytes).map_err(|e| e.to_string())?;
    decode(bytes).map_err(|e| e.to_string())
}

#[cfg(test)]
mod tests {
    use super::*;
    use vybe_ast::canon::ComponentBinaryDef;

    #[test]
    fn checked_decoder_rejects_truncated_sections_and_unrepresented_sections() {
        for suffix in [&[0, 0x80][..], &[0, 2, 3], &[7, 2, 2, 0x7f], &[9, 1, 0]] {
            let mut bytes = crate::component_format::HEADER.to_vec();
            bytes.extend_from_slice(suffix);
            assert!(decode(&bytes).is_err(), "{suffix:?}");
        }
    }

    #[test]
    fn attributed_name_without_attributes_preserves_plain_name() {
        let mut cursor = Cursor::new(&[2, 1, b'x', 0]);
        assert_eq!(external_name(&mut cursor).unwrap(), "x");
        assert_eq!(cursor.remaining(), 0);
        assert!(external_name(&mut Cursor::new(&[2, 1, b'x', 1])).is_err());
    }

    #[test]
    fn primitive_types_use_the_same_wire_facts_as_the_writer() {
        for byte in 0..=u8::MAX {
            if let Some(expected) = crate::component_format::primitive_name(byte) {
                let bytes = [byte];
                let mut cursor = Cursor::new(&bytes);
                assert_eq!(
                    valtype(&mut cursor).unwrap(),
                    ValSpec::Prim(expected.into())
                );
                assert_eq!(cursor.remaining(), 0);
            }
        }
    }

    fn write_leb(mut value: u32, out: &mut Vec<u8>) {
        loop {
            let mut byte = (value & 0x7f) as u8;
            value >>= 7;
            if value != 0 {
                byte |= 0x80;
            }
            out.push(byte);
            if value == 0 {
                break;
            }
        }
    }

    fn write_section(out: &mut Vec<u8>, id: u8, data: &[u8]) {
        out.push(id);
        write_leb(data.len() as u32, out);
        out.extend_from_slice(data);
    }

    #[test]
    fn custom_section_keeps_name_and_data() {
        let component = ComponentBinary {
            definitions: vec![ComponentBinaryDef::Custom {
                name: "build-info".into(),
                data: vec![0, 1, 0xff],
            }],
        };
        let bytes = crate::write_component_binary(&component).unwrap();
        assert_eq!(read_component_binary(&bytes).unwrap(), component);
    }

    #[test]
    fn nested_component_preserves_following_sibling_sections() {
        let nested = ComponentBinary {
            definitions: vec![ComponentBinaryDef::Custom {
                name: "inner".into(),
                data: vec![1],
            }],
        };
        let component = ComponentBinary {
            definitions: vec![
                ComponentBinaryDef::Custom {
                    name: "before".into(),
                    data: vec![2],
                },
                ComponentBinaryDef::NestedComponent(std::sync::Arc::new(nested)),
                ComponentBinaryDef::Custom {
                    name: "after".into(),
                    data: vec![3],
                },
            ],
        };
        let bytes = crate::write_component_binary(&component).unwrap();
        assert_eq!(read_component_binary(&bytes).unwrap(), component);
    }

    #[test]
    fn nested_core_function_body_is_validated() {
        let mut module = vec![0, 0x61, 0x73, 0x6d, 1, 0, 0, 0];
        module.extend_from_slice(&[1, 4, 1, 0x60, 0, 0]);
        module.extend_from_slice(&[3, 2, 1, 0]);
        module.extend_from_slice(&[10, 4, 1, 2, 0, 0x0b]);
        let mut component = vec![0, 0x61, 0x73, 0x6d, 0x0d, 0, 1, 0];
        component.extend_from_slice(&[1, module.len() as u8]);
        component.extend_from_slice(&module);
        let validate = |bytes: &[u8]| {
            wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
                .validate_all(bytes)
                .is_ok()
        };
        assert!(validate(&component));
        assert_eq!(
            read_component_binary(&component).unwrap().definitions.len(),
            1
        );

        *component.last_mut().unwrap() = 0;
        assert!(!validate(&component));
        assert!(
            read_component_binary(&component)
                .unwrap_err()
                .starts_with("invalid component binary:")
        );
    }

    #[test]
    fn many_nested_core_function_bodies_are_validated() {
        const COUNT: u32 = 1024;
        let mut module = vec![0, 0x61, 0x73, 0x6d, 1, 0, 0, 0];
        write_section(&mut module, 1, &[1, 0x60, 0, 0]);
        let mut functions = Vec::new();
        write_leb(COUNT, &mut functions);
        functions.resize(functions.len() + COUNT as usize, 0);
        write_section(&mut module, 3, &functions);
        let mut code = Vec::new();
        write_leb(COUNT, &mut code);
        for _ in 0..COUNT {
            code.extend_from_slice(&[2, 0, 0x0b]);
        }
        write_section(&mut module, 10, &code);
        let mut component = vec![0, 0x61, 0x73, 0x6d, 0x0d, 0, 1, 0];
        write_section(&mut component, 1, &module);
        assert_eq!(
            read_component_binary(&component).unwrap().definitions.len(),
            1
        );

        *component.last_mut().unwrap() = 0;
        assert!(read_component_binary(&component).is_err());
    }

    #[test]
    #[ignore = "manual component-reader benchmark"]
    fn benchmark_nested_component_validation() {
        const FUNCTIONS: u32 = 1024;
        const DEPTH: usize = 8;
        const RUNS: usize = 30;
        let mut module = vec![0, 0x61, 0x73, 0x6d, 1, 0, 0, 0];
        write_section(&mut module, 1, &[1, 0x60, 0, 0]);
        let mut functions = Vec::new();
        write_leb(FUNCTIONS, &mut functions);
        functions.resize(functions.len() + FUNCTIONS as usize, 0);
        write_section(&mut module, 3, &functions);
        let mut code = Vec::new();
        write_leb(FUNCTIONS, &mut code);
        for _ in 0..FUNCTIONS {
            code.extend_from_slice(&[2, 0, 0x0b]);
        }
        write_section(&mut module, 10, &code);
        let mut component = vec![0, 0x61, 0x73, 0x6d, 0x0d, 0, 1, 0];
        write_section(&mut component, 1, &module);
        for _ in 0..DEPTH {
            let mut parent = vec![0, 0x61, 0x73, 0x6d, 0x0d, 0, 1, 0];
            write_section(&mut parent, 4, &component);
            component = parent;
        }
        let start = std::time::Instant::now();
        for _ in 0..RUNS {
            std::hint::black_box(read_component_binary(&component).unwrap());
        }
        println!(
            "component reader: {DEPTH} nested components, {FUNCTIONS} funcs, \
             {RUNS} runs, {:.3} ms/run",
            start.elapsed().as_secs_f64() * 1000.0 / RUNS as f64
        );
    }
}
