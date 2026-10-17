//! Native data/element declarations decoded through the shared binary types.
use super::{WasmError, native_globals, section_types};
use crate::binary_cursor::Cursor;
use vybe_runtime::chunk::{
    ConstExpr, NativeElementSegment, NativeSegmentMode, NativeSegments, TypeEntry,
};
use vybe_runtime::opcode::heaptype::HeapType;

pub(super) fn read(
    data: &[u8],
    elements: &[u8],
    names: &[String],
    imports: usize,
    types: &[TypeEntry],
) -> Result<(Vec<Vec<u8>>, NativeSegments), WasmError> {
    let mut stack = Vec::new();
    let mut expression = |cursor: &mut Cursor<'_>| {
        native_globals::read_const_expr(cursor, names, imports, types, &mut stack)
    };
    let mut metadata = NativeSegments::default();
    let mut payloads = Vec::new();
    if !data.is_empty() {
        let mut cursor = Cursor::new(data);
        let count = section_types::count(&mut cursor, 1)?;
        metadata.data_modes.reserve(count);
        payloads.reserve(count);
        for _ in 0..count {
            let mode = match cursor.u32()? {
                0 => NativeSegmentMode::Active {
                    index: 0,
                    offset: expression(&mut cursor)?,
                },
                1 => NativeSegmentMode::Passive,
                2 => NativeSegmentMode::Active {
                    index: cursor.u32()?,
                    offset: expression(&mut cursor)?,
                },
                _ => return Err("invalid data segment flags".into()),
            };
            let length = cursor.u32()? as usize;
            payloads.push(cursor.bytes(length)?.to_vec());
            metadata.data_modes.push(mode);
        }
        section_types::finish(&cursor)?;
    }
    if !elements.is_empty() {
        let mut cursor = Cursor::new(elements);
        let count = section_types::count(&mut cursor, 1)?;
        metadata.elements.reserve(count);
        for _ in 0..count {
            let flags = cursor.u32()?;
            if flags > 7 {
                return Err("invalid element segment flags".into());
            }
            let mode = if flags & 1 == 0 {
                let index = if flags & 2 != 0 { cursor.u32()? } else { 0 };
                NativeSegmentMode::Active {
                    index,
                    offset: expression(&mut cursor)?,
                }
            } else if flags & 2 != 0 {
                NativeSegmentMode::Declarative
            } else {
                NativeSegmentMode::Passive
            };
            let reference_type = if flags & 4 != 0 {
                if flags == 4 {
                    (HeapType::Abstract(0x70), true, false, false)
                } else {
                    section_types::reference_type(&mut cursor)?
                }
            } else {
                if flags != 0 && cursor.byte()? != 0 {
                    return Err("invalid element kind".into());
                }
                (HeapType::Abstract(0x70), true, false, false)
            };
            let count = section_types::count(&mut cursor, 1)?;
            let mut items = Vec::with_capacity(count);
            for _ in 0..count {
                let item = if flags & 4 != 0 {
                    expression(&mut cursor)?
                } else {
                    let index = cursor.u32()?;
                    if (index as usize) < imports {
                        ConstExpr::ImportedRefFunc(index)
                    } else {
                        ConstExpr::RefFunc(index as usize - imports + 1)
                    }
                };
                items.push(item);
            }
            metadata.elements.push(NativeElementSegment {
                mode,
                reference_type,
                items: items.into(),
            });
        }
        section_types::finish(&cursor)?;
    }
    Ok((payloads, metadata))
}
