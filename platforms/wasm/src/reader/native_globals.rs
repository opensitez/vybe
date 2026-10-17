//! Preserve native global constant expressions for module instantiation.
use super::{WasmError, section_types};
use crate::binary_cursor::Cursor;
use vybe_runtime::Value;
use vybe_runtime::chunk::{ConstExpr, GlobalInit, NativeGlobalType, TypeEntry};

pub(super) fn read_imported_global_types(
    imports: &[u8],
) -> Result<Vec<NativeGlobalType>, WasmError> {
    let mut result = Vec::new();
    section_types::for_each_import(imports, |_, _, ty| {
        if let section_types::ImportType::Global(ty) = ty {
            result.push(ty.clone());
        }
        Ok(())
    })?;
    Ok(result)
}

pub(super) fn read_global_inits(
    bytes: &[u8],
    names: &[String],
    imported_functions: usize,
    types: &[TypeEntry],
    declarations: &mut Vec<NativeGlobalType>,
) -> Result<Vec<GlobalInit>, WasmError> {
    if bytes.is_empty() {
        return Ok(Vec::new());
    }
    let mut cursor = Cursor::new(bytes);
    let defined = section_types::count(&mut cursor, 1)?;
    let first = names
        .len()
        .checked_sub(defined)
        .ok_or_else(|| WasmError::invalid("native global index space is incomplete"))?;
    let mut inits = Vec::with_capacity(defined);
    declarations.reserve(defined);
    let mut stack = Vec::new();
    for offset in 0..defined {
        declarations.push(section_types::global_type(&mut cursor)?);
        inits.push(GlobalInit {
            name: names[first + offset].clone(),
            init: read_const_expr(
                &mut cursor,
                &names[..first + offset],
                imported_functions,
                types,
                &mut stack,
            )?,
        });
    }
    section_types::finish(&cursor)?;
    Ok(inits)
}

pub(super) fn read_const_expr(
    cursor: &mut Cursor<'_>,
    names: &[String],
    imported_functions: usize,
    types: &[TypeEntry],
    stack: &mut Vec<ConstExpr>,
) -> Result<ConstExpr, WasmError> {
    stack.clear();
    loop {
        let opcode = cursor.byte()?;
        let expression = match opcode {
            0x41 => ConstExpr::Value(Value::I32(cursor.i32()?)),
            0x42 => ConstExpr::Value(Value::I64(cursor.i64()?)),
            0x43 => ConstExpr::Value(Value::F32(f32::from_le_bytes(
                cursor.bytes(4)?.try_into().unwrap(),
            ))),
            0x44 => ConstExpr::Value(Value::F64(f64::from_le_bytes(
                cursor.bytes(8)?.try_into().unwrap(),
            ))),
            0xfd => {
                if cursor.u32()? != 12 {
                    return Err(WasmError::invalid("unsupported SIMD constant expression"));
                }
                ConstExpr::Value(Value::V128(cursor.bytes(16)?.try_into().unwrap()))
            }
            0xfb => match cursor.u32()? {
                0 => {
                    let index = cursor.u32()?;
                    let ty = types
                        .get(index as usize)
                        .filter(|ty| ty.kind == vybe_runtime::chunk::CompositeKind::Struct)
                        .ok_or_else(|| {
                            WasmError::invalid("global struct initializer type is not a struct")
                        })?;
                    let start = stack.len().checked_sub(ty.fields.len()).ok_or_else(|| {
                        WasmError::invalid("global struct initializer stack underflow")
                    })?;
                    let fields = stack.drain(start..).collect::<Vec<_>>().into();
                    ConstExpr::StructNew {
                        typeidx: super::native_type_index(index)?,
                        fields,
                    }
                }
                1 => {
                    let index = cursor.u32()?;
                    types
                        .get(index as usize)
                        .filter(|ty| ty.kind == vybe_runtime::chunk::CompositeKind::Struct)
                        .ok_or_else(|| {
                            WasmError::invalid(
                                "global default struct initializer type is not a struct",
                            )
                        })?;
                    ConstExpr::StructNewDefault {
                        typeidx: super::native_type_index(index)?,
                    }
                }
                28 => ConstExpr::RefI31(Box::new(stack.pop().ok_or_else(|| {
                    WasmError::invalid("global i31 initializer stack underflow")
                })?)),
                sub => {
                    return Err(WasmError::invalid(format!(
                        "unsupported native GC initializer: {sub:#x}"
                    )));
                }
            },
            0xd0 => {
                section_types::heap_type(cursor, true)?;
                ConstExpr::Value(Value::Null)
            }
            0x23 => {
                let index = cursor.u32()? as usize;
                let name = names
                    .get(index)
                    .ok_or_else(|| WasmError::invalid("global initializer index out of bounds"))?;
                ConstExpr::GlobalGet(name.clone())
            }
            0xd2 => {
                let index = cursor.u32()?;
                if (index as usize) < imported_functions {
                    ConstExpr::ImportedRefFunc(index)
                } else {
                    ConstExpr::RefFunc(index as usize - imported_functions + 1)
                }
            }
            0x6a | 0x7c => {
                let (left, right) = operands(stack)?;
                ConstExpr::Add(Box::new(left), Box::new(right))
            }
            0x6c | 0x7e => {
                let (left, right) = operands(stack)?;
                ConstExpr::Mul(Box::new(left), Box::new(right))
            }
            0x6b | 0x7d => {
                let minus_one = if opcode == 0x6b {
                    Value::I32(-1)
                } else {
                    Value::I64(-1)
                };
                let (left, right) = operands(stack)?;
                ConstExpr::Add(
                    Box::new(left),
                    Box::new(ConstExpr::Mul(
                        Box::new(right),
                        Box::new(ConstExpr::Value(minus_one)),
                    )),
                )
            }
            0x0b => break,
            other => {
                return Err(WasmError::invalid(format!(
                    "unsupported native global initializer: {other:#x}"
                )));
            }
        };
        stack.push(expression);
    }
    if stack.len() != 1 {
        return Err(WasmError::invalid(
            "global initializer must produce one value",
        ));
    }
    Ok(stack.pop().unwrap())
}

fn operands(stack: &mut Vec<ConstExpr>) -> Result<(ConstExpr, ConstExpr), WasmError> {
    let right = stack
        .pop()
        .ok_or_else(|| WasmError::invalid("global initializer stack underflow"))?;
    let left = stack
        .pop()
        .ok_or_else(|| WasmError::invalid("global initializer stack underflow"))?;
    Ok((left, right))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::encoding::*;

    fn module(globals: &[u8], imports: Option<&[u8]>) -> Vec<u8> {
        let mut bytes = Vec::from(WASM_MAGIC);
        bytes.extend_from_slice(&WASM_VERSION);
        write_section(&mut bytes, SECTION_TYPE, &[1, TYPE_FUNC, 0, 0]);
        if let Some(imports) = imports {
            write_section(&mut bytes, SECTION_IMPORT, imports);
        }
        write_section(&mut bytes, SECTION_FUNCTION, &[1, 0]);
        write_section(&mut bytes, SECTION_GLOBAL, globals);
        let mut exports = vec![1];
        write_name(&mut exports, "_start");
        exports.extend_from_slice(&[0, 0]);
        write_section(&mut bytes, SECTION_EXPORT, &exports);
        write_section(&mut bytes, SECTION_CODE, &[1, 2, 0, 0x0b]);
        bytes
    }

    fn decode(bytes: &[u8]) -> Vec<vybe_runtime::Chunk> {
        wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
            .validate_all(bytes)
            .expect("native initializer fixture validates independently");
        super::super::read_wasm(bytes).expect("native initializers decode")
    }

    fn struct_module(fields: &[u8], initializer: &[u8]) -> Vec<u8> {
        let mut bytes = Vec::from(WASM_MAGIC);
        bytes.extend_from_slice(&WASM_VERSION);
        let mut types = vec![2, 0x5f];
        types.extend_from_slice(fields);
        types.extend_from_slice(&[TYPE_FUNC, 0, 1, TYPE_I32]);
        write_section(&mut bytes, SECTION_TYPE, &types);
        write_section(&mut bytes, SECTION_FUNCTION, &[1, 1]);
        let mut globals = vec![1, 0x63, 0, 1];
        globals.extend_from_slice(initializer);
        globals.push(0x0b);
        write_section(&mut bytes, SECTION_GLOBAL, &globals);
        let mut exports = vec![1];
        write_name(&mut exports, "_start");
        exports.extend_from_slice(&[0, 0]);
        write_section(&mut bytes, SECTION_EXPORT, &exports);
        write_section(
            &mut bytes,
            SECTION_CODE,
            &[1, 8, 0, 0x23, 0, 0xfb, 2, 0, 0, 0x0b],
        );
        bytes
    }

    #[test]
    fn native_struct_global_is_typed_and_fresh_on_repeated_instantiation() {
        let chunks = decode(&struct_module(&[1, TYPE_I32, 1], &[0x41, 42, 0xfb, 0, 0]));
        let mut vm = vybe_runtime::VM::new();
        vm.run(chunks.clone()).unwrap();
        let first = vm.global("__wasm_instance_0_global_0").unwrap().clone();
        let Value::Object(first) = first else {
            panic!("typed struct global")
        };
        assert_ne!(first.lock().unwrap().type_id, 0);
        first.lock().unwrap().fields[0] = Value::I32(99);
        vm.run(chunks).unwrap();
        let second = vm.global("__wasm_instance_2_global_0").unwrap();
        let Value::Object(second) = second else {
            panic!("second typed struct global")
        };
        assert!(!std::sync::Arc::ptr_eq(&first, second));
        assert_eq!(first.lock().unwrap().fields[0], Value::I32(99));
        assert_eq!(second.lock().unwrap().fields[0], Value::I32(42));
    }

    #[test]
    fn struct_function_initializer_rebases_on_both_loaders() {
        let mut bytes = Vec::from(WASM_MAGIC);
        bytes.extend_from_slice(&WASM_VERSION);
        write_section(
            &mut bytes,
            SECTION_TYPE,
            &[2, TYPE_FUNC, 0, 1, TYPE_I32, 0x5f, 1, 0x64, 0, 0],
        );
        write_section(&mut bytes, SECTION_FUNCTION, &[2, 0, 0]);
        write_section(
            &mut bytes,
            SECTION_GLOBAL,
            &[1, 0x64, 1, 0, 0xd2, 0, 0xfb, 0, 1, 0x0b],
        );
        let mut exports = vec![1];
        write_name(&mut exports, "_start");
        exports.extend_from_slice(&[0, 1]);
        write_section(&mut bytes, SECTION_EXPORT, &exports);
        write_section(&mut bytes, 9, &[1, 3, 0, 1, 0]);
        write_section(
            &mut bytes,
            SECTION_CODE,
            &[
                2, 4, 0, 0x41, 42, 0x0b, 10, 0, 0x23, 0, 0xfb, 2, 1, 0, 0x14, 0, 0x0b,
            ],
        );
        let chunks = decode(&bytes);
        for linked in [false, true] {
            let mut vm = vybe_runtime::VM::new();
            for (instance, expected_function) in [(0, 1), (3, 4)] {
                let result = if linked {
                    vm.run_linked(chunks.clone(), vec![])
                } else {
                    vm.run(chunks.clone())
                };
                result.unwrap();
                let Value::Object(object) = vm
                    .global(&format!("__wasm_instance_{instance}_global_0"))
                    .unwrap()
                else {
                    panic!("struct global")
                };
                let object = object.lock().unwrap();
                let Value::Object(function) = &object.fields[0] else {
                    panic!("typed function field")
                };
                let function = function.lock().unwrap();
                let vybe_runtime::value::ObjectKind::Function(function) = &function.kind else {
                    panic!("callable function field")
                };
                assert_eq!(function.chunk_index, expected_function);
            }
        }
    }

    #[test]
    fn native_struct_default_global_has_typed_zero_fields() {
        let fields = [
            8, TYPE_I32, 1, TYPE_I64, 1, TYPE_F32, 1, TYPE_F64, 1, 0x7b, 1, 0x6e, 1, 0x78, 1, 0x77,
            1,
        ];
        let chunks = decode(&struct_module(&fields, &[0xfb, 1, 0]));
        let mut vm = vybe_runtime::VM::new();
        vm.run(chunks).unwrap();
        let Value::Object(object) = vm.global("__wasm_instance_0_global_0").unwrap() else {
            panic!("default struct global")
        };
        assert_eq!(
            object.lock().unwrap().fields,
            vec![
                Value::I32(0),
                Value::I64(0),
                Value::F32(0.0),
                Value::F64(0.0),
                Value::V128([0; 16]),
                Value::Null,
                Value::I32(0),
                Value::I32(0)
            ]
        );
    }

    #[test]
    fn numeric_initializers_are_owned_once_and_instantiated() {
        let mut globals = vec![4, 0x7f, 1, 0x41];
        write_leb128_i32(&mut globals, 65536);
        globals.extend_from_slice(&[0x0b, 0x7e, 0, 0x42]);
        write_leb128_i64(&mut globals, 4294967297);
        globals.extend_from_slice(&[0x0b, 0x7d, 0, 0x43]);
        globals.extend_from_slice(&1.5_f32.to_le_bytes());
        globals.extend_from_slice(&[0x0b, 0x7c, 0, 0x44]);
        globals.extend_from_slice(&(-2.25_f64).to_le_bytes());
        globals.push(0x0b);
        let chunks = decode(&module(&globals, None));
        assert_eq!(chunks[0].global_inits.len(), 4);
        assert!(chunks[1].global_inits.is_empty());
        let mut vm = vybe_runtime::VM::new();
        vm.run(chunks).unwrap();
        for (index, expected) in [
            Value::I32(65536),
            Value::I64(4294967297),
            Value::F32(1.5),
            Value::F64(-2.25),
        ]
        .into_iter()
        .enumerate()
        {
            assert_eq!(
                vm.global(&format!("__wasm_instance_0_global_{index}")),
                Some(&expected)
            );
        }
    }

    #[test]
    fn imported_global_initializers_resolve_the_declared_binding() {
        let mut imports = vec![1];
        write_name(&mut imports, "env");
        write_name(&mut imports, "seed");
        imports.extend_from_slice(&[3, 0x7f, 0]);
        let chunks = decode(&module(&[1, 0x7f, 0, 0x23, 0, 0x0b], Some(&imports)));
        let mut vm = vybe_runtime::VM::new();
        vm.set_global(
            &vybe_runtime::chunk::imported_global_key("env", "seed"),
            Value::I32(37),
        );
        vm.run(chunks).unwrap();
        assert_eq!(
            vm.global("__wasm_instance_0_global_1"),
            Some(&Value::I32(37))
        );
    }

    #[test]
    fn local_function_reference_initializers_are_rebased_after_prior_modules() {
        let chunks = decode(&module(
            &[2, 0x70, 0, 0xd2, 0, 0x0b, 0x6f, 0, 0xd0, 0x6f, 0x0b],
            None,
        ));
        let mut prefix = vybe_runtime::Chunk::new("prior");
        prefix.emit_op(vybe_runtime::Op::RETURN, 0);
        let mut vm = vybe_runtime::VM::new();
        vm.run(vec![prefix]).unwrap();
        vm.run(chunks).unwrap();
        let Value::Object(reference) = vm.global("__wasm_instance_1_global_0").unwrap() else {
            panic!("function initializer must retain a callable reference");
        };
        let object = reference.lock().unwrap();
        let vybe_runtime::value::ObjectKind::Function(function) = &object.kind else {
            panic!("function initializer must be callable");
        };
        assert_eq!(function.chunk_index, 2);
        assert_eq!(vm.global("__wasm_instance_1_global_1"), Some(&Value::Null));
    }

    #[test]
    fn extended_integer_initializers_keep_wrapping_arithmetic() {
        let mut globals = vec![1, 0x7f, 0, 0x41];
        write_leb128_i32(&mut globals, i32::MIN);
        globals.extend_from_slice(&[0x41, 1, 0x6b, 0x41, 2, 0x6c, 0x0b]);
        let mut vm = vybe_runtime::VM::new();
        vm.run(decode(&module(&globals, None))).unwrap();
        assert_eq!(
            vm.global("__wasm_instance_0_global_0"),
            Some(&Value::I32(-2))
        );
    }

    #[test]
    fn simd_global_initializers_preserve_all_bits() {
        let payload = std::array::from_fn::<_, 16, _>(|index| (index * 17) as u8);
        let mut globals = vec![1, 0x7b, 0, 0xfd, 0x0c];
        globals.extend_from_slice(&payload);
        globals.push(0x0b);
        let mut vm = vybe_runtime::VM::new();
        vm.run(decode(&module(&globals, None))).unwrap();
        assert_eq!(
            vm.global("__wasm_instance_0_global_0"),
            Some(&Value::V128(payload))
        );
    }
}
