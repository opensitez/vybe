//! Native recursive types decoded in-house, retaining structural identity.
use super::WasmError;
use crate::binary_cursor::Cursor;
use std::collections::HashMap;
use vybe_runtime::chunk::{CompositeKind, TypeEntry};
use vybe_runtime::native_gc as native;

struct GroupContext<'a> {
    start: u32,
    len: u32,
    previous: &'a [TypeEntry],
    dependencies: Vec<native::TypeIdentity>,
    dependency_indices: HashMap<native::TypeIdentity, u32>,
}
fn count(cursor: &mut Cursor<'_>, minimum_bytes: usize) -> Result<usize, WasmError> {
    let n = cursor.u32()? as usize;
    if n > cursor.remaining() / minimum_bytes {
        return Err("truncated native type vector".into());
    }
    Ok(n)
}
fn table_index(index: u32) -> Result<u16, WasmError> {
    index
        .checked_add(1)
        .and_then(|n| u16::try_from(n).ok())
        .ok_or_else(|| WasmError::invalid("native type index exceeds supported range"))
}
fn abstract_type(byte: u8) -> Option<(u8, &'static str, &'static str)> {
    // Identity tags are retained independently of the binary spelling.
    Some(match byte {
        0x70 => (0, "func", "func"),
        0x6f => (1, "extern", "extern"),
        0x6e => (2, "any", "any"),
        0x71 => (3, "none", "null"),
        0x72 => (4, "noextern", "nullextern"),
        0x73 => (5, "nofunc", "nullfunc"),
        0x6d => (6, "eq", "eq"),
        0x6b => (7, "struct", "struct"),
        0x6a => (8, "array", "array"),
        0x6c => (9, "i31", "i31"),
        0x69 => (10, "exn", "exn"),
        0x74 => (11, "noexn", "nullexn"),
        0x68 => (12, "cont", "cont"),
        0x75 => (13, "nocont", "nullcont"),
        _ => return None,
    })
}
impl GroupContext<'_> {
    fn reference(&mut self, index: u32) -> Result<native::TypeRef, WasmError> {
        if index >= self.start && index - self.start < self.len {
            return Ok(native::TypeRef::Local(index - self.start));
        }
        let identity = self
            .previous
            .get(index as usize)
            .and_then(|entry| entry.native_identity.as_ref())
            .ok_or_else(|| {
                WasmError::invalid("native type reference is outside its recursive group")
            })?;
        let next = self.dependencies.len() as u32;
        let dependency = *self
            .dependency_indices
            .entry(identity.clone())
            .or_insert_with(|| {
                self.dependencies.push(identity.clone());
                next
            });
        Ok(native::TypeRef::Dependency(dependency))
    }
    fn value(&mut self, cursor: &mut Cursor<'_>) -> Result<(native::ValueType, String), WasmError> {
        use vybe_runtime::chunk::NativeValueType as V;
        use vybe_runtime::opcode::heaptype::HeapType as H;
        let value = super::section_types::value_type(cursor)?;
        let scalar = match value {
            V::I32 => Some((0x7f, "i32")),
            V::I64 => Some((0x7e, "i64")),
            V::F32 => Some((0x7d, "f32")),
            V::F64 => Some((0x7c, "f64")),
            V::V128 => Some((0x7b, "v128")),
            _ => None,
        };
        if let Some((byte, name)) = scalar {
            return Ok((native::ValueType::Scalar(byte), name.into()));
        }
        let V::Ref((heap, nullable, exact, shared)) = value else {
            unreachable!()
        };
        let (heap, label) = match heap {
            H::Concrete(index) => {
                let index = index
                    .checked_sub(1)
                    .ok_or_else(|| WasmError::invalid("reserved native type index"))?;
                let label = if exact {
                    format!("(ref (exact (module {index})))")
                } else if nullable {
                    format!("(ref null (module {index}))")
                } else {
                    format!("(ref (module {index}))")
                };
                (native::HeapType::Concrete(self.reference(index)?), label)
            }
            H::Abstract(byte) => {
                let (kind, full, short) =
                    abstract_type(byte).ok_or("invalid abstract heap type")?;
                let label = match (nullable, shared) {
                    (true, false) => format!("{short}ref"),
                    (true, true) => format!("(shared {short}ref)"),
                    (false, false) => format!("(ref {full})"),
                    (false, true) => format!("(ref (shared {full}))"),
                };
                (native::HeapType::Abstract { kind, shared }, label)
            }
        };
        Ok((
            native::ValueType::Ref {
                nullable,
                exact,
                heap,
            },
            label,
        ))
    }
    fn field(&mut self, cursor: &mut Cursor<'_>) -> Result<(native::FieldType, String), WasmError> {
        let (storage, label) = match cursor.peek() {
            Some(0x78) => {
                cursor.byte()?;
                (native::StorageType::I8, "i8".into())
            }
            Some(0x77) => {
                cursor.byte()?;
                (native::StorageType::I16, "i16".into())
            }
            _ => {
                let (value, label) = self.value(cursor)?;
                (native::StorageType::Value(value), label)
            }
        };
        let mutable = match cursor.byte()? {
            0 => false,
            1 => true,
            _ => return Err("invalid field mutability".into()),
        };
        Ok((native::FieldType { storage, mutable }, label))
    }
    fn values(
        &mut self,
        cursor: &mut Cursor<'_>,
    ) -> Result<(Vec<native::ValueType>, String), WasmError> {
        let length = count(cursor, 1)?;
        let mut values = Vec::with_capacity(length);
        let mut names = Vec::with_capacity(length);
        for _ in 0..length {
            let (value, name) = self.value(cursor)?;
            values.push(value);
            names.push(name);
        }
        Ok((values, names.join(",")))
    }
    fn subtype(
        &mut self,
        cursor: &mut Cursor<'_>,
    ) -> Result<(native::SubType, TypeEntry), WasmError> {
        let mut final_ = true;
        let mut supertypes = Vec::new();
        let mut parent_index = 0;
        if matches!(cursor.peek(), Some(0x4f | 0x50)) {
            final_ = cursor.byte()? == 0x4f;
            let length = count(cursor, 1)?;
            if length > 1 {
                return Err(WasmError::invalid("multiple native supertypes"));
            }
            for _ in 0..length {
                let index = cursor.u32()?;
                parent_index = table_index(index)?;
                supertypes.push(self.reference(index)?);
            }
        }
        let shared = cursor.peek() == Some(0x65);
        if shared {
            cursor.byte()?;
        }
        let mut describes_index = 0;
        let describes = if cursor.peek() == Some(0x4c) {
            cursor.byte()?;
            let index = cursor.u32()?;
            describes_index = table_index(index)?;
            Some(self.reference(index)?)
        } else {
            None
        };
        let descriptor = if cursor.peek() == Some(0x4d) {
            cursor.byte()?;
            Some(self.reference(cursor.u32()?)?)
        } else {
            None
        };
        let (composite, kind, fields) = match cursor.byte()? {
            0x60 => {
                let (params, param_names) = self.values(cursor)?;
                let (results, result_names) = self.values(cursor)?;
                (
                    native::CompositeType::Func { params, results },
                    CompositeKind::Func,
                    vec![param_names, result_names],
                )
            }
            0x5f => {
                let length = count(cursor, 2)?;
                let mut fields = Vec::with_capacity(length);
                for _ in 0..length {
                    fields.push(self.field(cursor)?.0);
                }
                let names = (0..length).map(|i| i.to_string()).collect();
                (
                    native::CompositeType::Struct(fields),
                    CompositeKind::Struct,
                    names,
                )
            }
            0x5e => {
                let (field, name) = self.field(cursor)?;
                (
                    native::CompositeType::Array(field),
                    CompositeKind::Array,
                    vec![name],
                )
            }
            0x5d => (
                native::CompositeType::Cont(self.reference(cursor.u32()?)?),
                CompositeKind::Cont,
                vec![],
            ),
            _ => return Err("invalid native composite type".into()),
        };
        if (descriptor.is_some() || describes.is_some()) && kind != CompositeKind::Struct {
            return Err(WasmError::invalid(
                "custom descriptors require struct types",
            ));
        }
        Ok((
            native::SubType {
                final_,
                shared,
                supertypes,
                descriptor,
                describes,
                composite,
            },
            TypeEntry {
                kind,
                fields,
                parent_index,
                describes_index,
                ..Default::default()
            },
        ))
    }
}

pub(super) fn read_type_entries(bytes: &[u8]) -> Result<Vec<TypeEntry>, WasmError> {
    if bytes.is_empty() {
        return Ok(Vec::new());
    }
    let mut cursor = Cursor::new(bytes);
    let groups = count(&mut cursor, 1)?;
    let mut entries = Vec::new();
    for _ in 0..groups {
        let length = if cursor.peek() == Some(0x4e) {
            cursor.byte()?;
            count(&mut cursor, 1)?
        } else {
            1
        };
        let start = u32::try_from(entries.len())
            .map_err(|_| WasmError::invalid("native type count overflow"))?;
        let mut context = GroupContext {
            start,
            len: length as u32,
            previous: &entries,
            dependencies: Vec::new(),
            dependency_indices: HashMap::new(),
        };
        let mut shapes = Vec::with_capacity(length);
        let mut rows = Vec::with_capacity(length);
        for _ in 0..length {
            let (shape, row) = context.subtype(&mut cursor)?;
            shapes.push(shape);
            rows.push(row);
        }
        let identity = native::RecGroup::new(shapes, context.dependencies);
        for (position, mut row) in rows.into_iter().enumerate() {
            row.name = format!("native-type-{}", entries.len());
            row.native_identity = identity.identity(position as u32);
            entries.push(row);
        }
    }
    if cursor.remaining() != 0 {
        return Err("trailing bytes in native type section".into());
    }
    Ok(entries)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn malformed_types_and_padded_abstract_heaps_are_rejected() {
        for bytes in [
            &[1, 0x60, 0][..],
            &[1, 0x5f, 1, 0x7f],
            &[1, 0x5e, 0x7f, 2],
            &[1, 0x60, 1, 0x63, 0x80],
            &[1, 0x60, 1, 0x63, 0xef, 0x7f, 0],
            &[1, 0x60, 0, 0, 0],
        ] {
            assert!(read_type_entries(bytes).is_err(), "{bytes:?}");
        }
    }

    #[test]
    fn packed_and_exact_fields_keep_their_native_storage_identity() {
        let entries =
            read_type_entries(&[1, 0x4e, 1, 0x5f, 3, 0x78, 1, 0x77, 0, 0x63, 0x62, 0, 1]).unwrap();
        let identity = entries[0].native_identity.as_ref().unwrap();
        let native::CompositeType::Struct(fields) = &identity.definition().composite else {
            panic!("struct")
        };
        assert_eq!(fields[0].storage, native::StorageType::I8);
        assert!(fields[0].mutable);
        assert_eq!(fields[1].storage, native::StorageType::I16);
        assert!(!fields[1].mutable);
        assert_eq!(
            fields[2].storage,
            native::StorageType::Value(native::ValueType::Ref {
                nullable: true,
                exact: true,
                heap: native::HeapType::Concrete(native::TypeRef::Local(0))
            })
        );
    }

    #[test]
    fn descriptor_links_and_shared_reference_types_are_retained() {
        let entries = read_type_entries(&[
            1, 0x4e, 2, 0x4d, 1, 0x5f, 1, 0x63, 0x65, 0x6f, 0, 0x4c, 0, 0x5f, 0,
        ])
        .unwrap();
        let a = entries[0].native_identity.as_ref().unwrap().definition();
        let b = entries[1].native_identity.as_ref().unwrap().definition();
        assert_eq!(a.descriptor, Some(native::TypeRef::Local(1)));
        assert_eq!(b.describes, Some(native::TypeRef::Local(0)));
        assert_eq!(entries[1].describes_index, 1);
        let native::CompositeType::Struct(fields) = &a.composite else {
            panic!("struct")
        };
        assert_eq!(
            fields[0].storage,
            native::StorageType::Value(native::ValueType::Ref {
                nullable: true,
                exact: false,
                heap: native::HeapType::Abstract {
                    kind: 1,
                    shared: true
                }
            })
        );
    }

    #[test]
    fn recursive_group_identity_ignores_module_indices_and_names() {
        let mut a = read_type_entries(&[1, 0x4e, 1, 0x5f, 1, 0x63, 0, 0]).unwrap();
        let mut b = read_type_entries(&[2, 0x60, 0, 0, 0x4e, 1, 0x5f, 1, 0x63, 1, 0]).unwrap();
        assert_eq!(a[0].native_identity, b[1].native_identity);
        a[0].name = "first module".into();
        b[1].name = "second module".into();
        let mut registry = vybe_runtime::typedef::TypeRegistry::new();
        registry.load_type_table(&a);
        registry.load_type_table(&b);
        assert_eq!(registry.get_entry_id(&a[0]), registry.get_entry_id(&b[1]));
        assert!(registry.get_id("first module").is_none());
        assert!(registry.get_id("second module").is_none());
    }

    #[test]
    fn same_diagnostic_names_do_not_merge_different_native_shapes() {
        let a = read_type_entries(&[1, 0x5f, 1, 0x7f, 0]).unwrap();
        let b = read_type_entries(&[1, 0x5f, 1, 0x7e, 0]).unwrap();
        assert_eq!(a[0].name, b[0].name);
        let mut registry = vybe_runtime::typedef::TypeRegistry::new();
        registry.load_type_table(&a);
        registry.load_type_table(&b);
        assert_ne!(registry.get_entry_id(&a[0]), registry.get_entry_id(&b[0]));
        let language_type = TypeEntry {
            name: a[0].name.clone(),
            ..Default::default()
        };
        registry.load_type_table(std::slice::from_ref(&language_type));
        assert_ne!(
            registry.get_entry_id(&a[0]),
            registry.get_entry_id(&language_type)
        );
    }
}
