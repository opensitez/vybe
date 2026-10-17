//! Shared binary types for core section readers. No owned names or payloads.
use super::WasmError;
use crate::binary_cursor::Cursor;
use vybe_runtime::chunk::{NativeGlobalType, NativeValueType};
use vybe_runtime::opcode::heaptype::HeapType;

pub(super) type ReferenceType = (HeapType, bool, bool, bool);

#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) struct TableType {
    pub element_type: ReferenceType,
    pub initial: u64,
    pub maximum: Option<u64>,
    pub shared: bool,
    pub table64: bool,
}
#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) struct MemoryType {
    pub initial: u64,
    pub maximum: Option<u64>,
    pub shared: bool,
    pub memory64: bool,
    pub page_size_log2: Option<u32>,
}
#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) enum ImportType {
    Func { type_index: u32, exact: bool },
    Table(TableType),
    Memory(MemoryType),
    Global(NativeGlobalType),
    Tag(u32),
}
pub(super) fn count(cursor: &mut Cursor<'_>, minimum_bytes: usize) -> Result<usize, WasmError> {
    let count = cursor.u32()? as usize;
    if count > cursor.remaining() / minimum_bytes {
        return Err("truncated section vector".into());
    }
    Ok(count)
}
pub(super) fn finish(cursor: &Cursor<'_>) -> Result<(), WasmError> {
    if cursor.remaining() != 0 {
        return Err("trailing section bytes".into());
    }
    Ok(())
}
pub(super) fn reference_type(cursor: &mut Cursor<'_>) -> Result<ReferenceType, WasmError> {
    let nullable = match cursor.peek() {
        Some(0x63) => {
            cursor.byte()?;
            true
        }
        Some(0x64) => {
            cursor.byte()?;
            false
        }
        Some(0x65 | 0x68..=0x75) => true,
        _ => return Err("invalid reference type".into()),
    };
    heap_type(cursor, nullable)
}
pub(super) fn heap_type(
    cursor: &mut Cursor<'_>,
    nullable: bool,
) -> Result<ReferenceType, WasmError> {
    let shared = cursor.peek() == Some(0x65);
    if shared {
        cursor.byte()?;
    }
    let exact = cursor.peek() == Some(0x62);
    if exact {
        cursor.byte()?;
    }
    let start = cursor.position();
    let value = if exact {
        i64::from(cursor.u32()?)
    } else {
        cursor.s33()?
    };
    let heap = if value >= 0 {
        if shared {
            return Err("shared concrete heap type encoding".into());
        }
        let index = u32::try_from(value)
            .ok()
            .and_then(|n| n.checked_add(1))
            .ok_or_else(|| WasmError::invalid("native heap type index overflow"))?;
        HeapType::Concrete(index)
    } else {
        let byte = (value & 0x7f) as u8;
        if cursor.position() - start != 1 || !matches!(byte, 0x68..=0x75) {
            return Err("invalid abstract heap type encoding".into());
        }
        HeapType::Abstract(byte)
    };
    Ok((heap, nullable, exact, shared))
}
pub(super) fn value_type(cursor: &mut Cursor<'_>) -> Result<NativeValueType, WasmError> {
    let scalar = match cursor.peek() {
        Some(0x7f) => Some(NativeValueType::I32),
        Some(0x7e) => Some(NativeValueType::I64),
        Some(0x7d) => Some(NativeValueType::F32),
        Some(0x7c) => Some(NativeValueType::F64),
        Some(0x7b) => Some(NativeValueType::V128),
        _ => None,
    };
    if let Some(value) = scalar {
        cursor.byte()?;
        return Ok(value);
    }
    match cursor.peek() {
        Some(0x63..=0x65 | 0x68..=0x75) => Ok(NativeValueType::Ref(reference_type(cursor)?)),
        _ => Err("invalid value type".into()),
    }
}
pub(super) fn global_type(cursor: &mut Cursor<'_>) -> Result<NativeGlobalType, WasmError> {
    let value_type = value_type(cursor)?;
    let flags = cursor.byte()?;
    if flags & !3 != 0 {
        return Err("invalid global type flags".into());
    }
    Ok(NativeGlobalType {
        value_type,
        mutable: flags & 1 != 0,
        shared: flags & 2 != 0,
    })
}
fn limits(cursor: &mut Cursor<'_>, flags: u8) -> Result<(u64, Option<u64>), WasmError> {
    let mut read = || -> Result<u64, WasmError> {
        Ok(if flags & 4 != 0 {
            cursor.u64()?
        } else {
            u64::from(cursor.u32()?)
        })
    };
    let initial = read()?;
    let maximum = if flags & 1 != 0 { Some(read()?) } else { None };
    Ok((initial, maximum))
}
pub(super) fn table_type(cursor: &mut Cursor<'_>) -> Result<TableType, WasmError> {
    let element_type = reference_type(cursor)?;
    let flags = cursor.byte()?;
    if flags & !7 != 0 {
        return Err("invalid table type flags".into());
    }
    let (initial, maximum) = limits(cursor, flags)?;
    Ok(TableType {
        element_type,
        initial,
        maximum,
        shared: flags & 2 != 0,
        table64: flags & 4 != 0,
    })
}
pub(super) fn memory_type(cursor: &mut Cursor<'_>) -> Result<MemoryType, WasmError> {
    let flags = cursor.byte()?;
    if flags & !15 != 0 {
        return Err("invalid memory type flags".into());
    }
    let (initial, maximum) = limits(cursor, flags)?;
    let page_size_log2 = if flags & 8 != 0 {
        Some(cursor.u32()?)
    } else {
        None
    };
    Ok(MemoryType {
        initial,
        maximum,
        shared: flags & 2 != 0,
        memory64: flags & 4 != 0,
        page_size_log2,
    })
}
fn import_type(cursor: &mut Cursor<'_>) -> Result<ImportType, WasmError> {
    Ok(match cursor.byte()? {
        kind @ (0 | 0x20) => ImportType::Func {
            type_index: cursor.u32()?,
            exact: kind == 0x20,
        },
        1 => ImportType::Table(table_type(cursor)?),
        2 => ImportType::Memory(memory_type(cursor)?),
        3 => ImportType::Global(global_type(cursor)?),
        4 => {
            if cursor.byte()? != 0 {
                return Err("invalid tag attribute".into());
            }
            ImportType::Tag(cursor.u32()?)
        }
        _ => return Err("invalid import kind".into()),
    })
}
pub(super) fn for_each_import<'a>(
    bytes: &'a [u8],
    mut visit: impl FnMut(&'a str, &'a str, &ImportType) -> Result<(), WasmError>,
) -> Result<(), WasmError> {
    if bytes.is_empty() {
        return Ok(());
    }
    let mut cursor = Cursor::new(bytes);
    let groups = count(&mut cursor, 1)?;
    for _ in 0..groups {
        let module = cursor.name()?;
        let name = cursor.name()?;
        match (name, cursor.peek()) {
            ("", Some(0x7f)) => {
                cursor.byte()?;
                let length = count(&mut cursor, 1)?;
                for _ in 0..length {
                    let name = cursor.name()?;
                    let ty = import_type(&mut cursor)?;
                    visit(module, name, &ty)?;
                }
            }
            ("", Some(0x7e)) => {
                cursor.byte()?;
                let ty = import_type(&mut cursor)?;
                let length = count(&mut cursor, 1)?;
                for _ in 0..length {
                    visit(module, cursor.name()?, &ty)?;
                }
            }
            _ => visit(module, name, &import_type(&mut cursor)?)?,
        }
    }
    finish(&cursor)
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn compact_imports_expand_without_name_or_type_clones() {
        let bytes = [
            2, 1, b'm', 0, 0x7f, 2, 1, b'a', 0, 0, 1, b'b', 3, 0x7f, 1, 1, b'n', 0, 0x7e, 0, 2, 2,
            1, b'c', 1, b'd',
        ];
        let mut rows = Vec::new();
        for_each_import(&bytes, |module, name, ty| {
            rows.push((module, name, ty.clone()));
            Ok(())
        })
        .unwrap();
        assert_eq!(rows.len(), 4);
        assert_eq!(rows[0].0, "m");
        assert_eq!(rows[3].1, "d");
        assert!(matches!(
            rows[1].2,
            ImportType::Global(NativeGlobalType { mutable: true, .. })
        ));
        assert_eq!(rows[2].2, rows[3].2);
    }
    #[test]
    fn shared_exact_and_wide_limits_have_distinct_facts() {
        let mut cursor = Cursor::new(&[0x63, 0x65, 0x6f]);
        assert_eq!(
            reference_type(&mut cursor).unwrap(),
            (HeapType::Abstract(0x6f), true, false, true)
        );
        let mut cursor = Cursor::new(&[0x64, 0x62, 0x80, 1]);
        assert_eq!(
            reference_type(&mut cursor).unwrap(),
            (HeapType::Concrete(129), false, true, false)
        );
        let mut cursor = Cursor::new(&[
            0x70, 5, 0x80, 0x80, 0x80, 0x80, 0x10, 0x81, 0x80, 0x80, 0x80, 0x10,
        ]);
        let table = table_type(&mut cursor).unwrap();
        assert_eq!(table.initial, 1u64 << 32);
        assert_eq!(table.maximum, Some((1u64 << 32) + 1));
        assert!(table.table64);
    }
    #[test]
    fn malformed_flags_names_and_trailing_payloads_fail() {
        assert!(global_type(&mut Cursor::new(&[0x7f, 4])).is_err());
        assert!(table_type(&mut Cursor::new(&[0x70, 8, 0])).is_err());
        assert!(reference_type(&mut Cursor::new(&[0x63, 0xef, 0x7f])).is_err());
        for bytes in [&[0, 0][..], &[1, 1, 0xff, 0, 0, 0], &[1, 0, 0, 0x7f, 1, 0]] {
            assert!(for_each_import(bytes, |_, _, _| Ok(())).is_err());
        }
    }
}
