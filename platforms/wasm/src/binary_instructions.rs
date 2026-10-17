//! Borrowing instruction-boundary decoding for native writer passes.
//! This decodes immediates; module typing remains the validator's responsibility.
use crate::binary_cursor::{Cursor, DecodeError};

#[derive(Debug, Clone, Copy)]
pub(crate) struct Instruction {
    pub start: usize,
    pub end: usize,
    pub opcode: u8,
    pub index: Option<u32>,
}

fn error(cursor: &Cursor<'_>, reason: &'static str) -> DecodeError {
    DecodeError {
        offset: cursor.position(),
        reason,
    }
}
fn heap(cursor: &mut Cursor<'_>) -> Result<(), DecodeError> {
    // Exact heap types carry a u32, unlike ordinary signed-33 heap types.
    let mut first = cursor.byte()?;
    if first == 0x65 {
        first = cursor.byte()?;
    }
    if first == 0x62 {
        cursor.u32()?;
        return Ok(());
    }
    let mut bytes = [0u8; 5];
    bytes[0] = first;
    let mut count = 1;
    while bytes[count - 1] & 0x80 != 0 {
        if count == bytes.len() {
            return Err(error(cursor, "heap type representation too long"));
        }
        bytes[count] = cursor.byte()?;
        count += 1;
    }
    let value = Cursor::new(&bytes[..count]).s33()?;
    if value < 0 && (count != 1 || !matches!((value & 0x7f) as u8, 0x66..=0x75)) {
        return Err(error(cursor, "invalid abstract heap type"));
    }
    Ok(())
}
fn block_type(cursor: &mut Cursor<'_>) -> Result<(), DecodeError> {
    match cursor.peek() {
        Some(0x63 | 0x64 | 0x65) => value_type(cursor),
        _ => {
            cursor.s33()?;
            Ok(())
        }
    }
}
pub(crate) fn value_type(cursor: &mut Cursor<'_>) -> Result<(), DecodeError> {
    match cursor.byte()? {
        0x63 | 0x64 => heap(cursor),
        0x65 => heap(cursor),
        0x66..=0x75 | 0x7b..=0x7f => Ok(()),
        _ => Err(error(cursor, "invalid value type")),
    }
}
fn memarg(cursor: &mut Cursor<'_>) -> Result<(), DecodeError> {
    let flags = cursor.u32()?;
    if flags & 0x40 != 0 {
        cursor.u32()?;
    }
    cursor.u64()?;
    Ok(())
}
fn indices(cursor: &mut Cursor<'_>, count: usize) -> Result<(), DecodeError> {
    for _ in 0..count {
        cursor.u32()?;
    }
    Ok(())
}
fn handlers(cursor: &mut Cursor<'_>) -> Result<(), DecodeError> {
    let count = cursor.u32()?;
    for _ in 0..count {
        match cursor.byte()? {
            0 => indices(cursor, 2)?,
            1 => indices(cursor, 1)?,
            _ => return Err(error(cursor, "invalid continuation handler")),
        }
    }
    Ok(())
}
pub(crate) fn simd_arity(sub: u32) -> Option<(usize, usize)> {
    let inputs = match sub {
        0x0c => 0,
        0x00..=0x0a
        | 0x0f..=0x16
        | 0x18..=0x19
        | 0x1b
        | 0x1d
        | 0x1f
        | 0x21
        | 0x4d
        | 0x53
        | 0x5c..=0x64
        | 0x67..=0x6a
        | 0x74..=0x75
        | 0x7a
        | 0x7c..=0x81
        | 0x83..=0x84
        | 0x87..=0x8a
        | 0x94
        | 0xa0..=0xa1
        | 0xa3..=0xa4
        | 0xa7..=0xaa
        | 0xc0..=0xc1
        | 0xc3..=0xc4
        | 0xc7..=0xca
        | 0xe0..=0xe1
        | 0xe3
        | 0xec..=0xed
        | 0xef
        | 0xf8..=0xff
        | 0x101..=0x104 => 1,
        0x52 | 0x105..=0x10c | 0x113 => 3,
        0x0b
        | 0x0d..=0x0e
        | 0x17
        | 0x1a
        | 0x1c
        | 0x1e
        | 0x20
        | 0x22..=0x4c
        | 0x4e..=0x51
        | 0x54..=0x5b
        | 0x65..=0x66
        | 0x6b..=0x73
        | 0x76..=0x79
        | 0x7b
        | 0x82
        | 0x85..=0x86
        | 0x8b..=0x93
        | 0x95..=0x99
        | 0x9b..=0x9f
        | 0xab..=0xae
        | 0xb1
        | 0xb5..=0xba
        | 0xbc..=0xbf
        | 0xcb..=0xce
        | 0xd1
        | 0xd5..=0xdf
        | 0xe4..=0xeb
        | 0xf0..=0xf7
        | 0x100
        | 0x10d..=0x112 => 2,
        _ => return None,
    };
    Some((inputs, usize::from(!matches!(sub, 0x0b | 0x58..=0x5b))))
}
pub(crate) fn simd(cursor: &mut Cursor<'_>) -> Result<u32, DecodeError> {
    let sub = cursor.u32()?;
    if simd_arity(sub).is_none() {
        return Err(error(cursor, "unknown SIMD instruction"));
    }
    match sub {
        0..=11 | 92 | 93 => memarg(cursor)?,
        12 | 13 => {
            cursor.bytes(16)?;
        }
        21..=34 => {
            cursor.byte()?;
        }
        84..=91 => {
            memarg(cursor)?;
            cursor.byte()?;
        }
        _ => {}
    }
    Ok(sub)
}
pub(crate) fn next(cursor: &mut Cursor<'_>) -> Result<Instruction, DecodeError> {
    let start = cursor.position();
    let opcode = cursor.byte()?;
    let mut index = None;
    match opcode {
        0x00
        | 0x01
        | 0x05
        | 0x0a
        | 0x0b
        | 0x0f
        | 0x19
        | 0x1a
        | 0x1b
        | 0x45..=0xc4
        | 0xd1
        | 0xd3
        | 0xd4 => {}
        0x02..=0x04 | 0x06 => block_type(cursor)?,
        0x07..=0x09
        | 0x0c..=0x0d
        | 0x10
        | 0x12
        | 0x14
        | 0x15
        | 0x18
        | 0x25
        | 0x26
        | 0x3f
        | 0x40
        | 0xd2
        | 0xd5
        | 0xd6
        | 0xe0
        | 0xe2 => {
            cursor.u32()?;
        }
        0x20..=0x24 => {
            index = Some(cursor.u32()?);
        }
        0x0e => {
            let count = cursor.u32()?;
            // Every label occupies at least one byte, including the default.
            if u64::from(count) + 1 > cursor.remaining() as u64 {
                return Err(error(cursor, "truncated branch table"));
            }
            for _ in 0..=count {
                cursor.u32()?;
            }
        }
        0x11 | 0x13 | 0xe1 | 0xe6 => indices(cursor, 2)?,
        0x1c => {
            let count = cursor.u32()?;
            for _ in 0..count {
                value_type(cursor)?;
            }
        }
        0x1f => {
            block_type(cursor)?;
            let count = cursor.u32()?;
            for _ in 0..count {
                match cursor.byte()? {
                    0 | 1 => indices(cursor, 2)?,
                    2 | 3 => indices(cursor, 1)?,
                    _ => return Err(error(cursor, "invalid try_table catch kind")),
                }
            }
        }
        0x28..=0x3e => memarg(cursor)?,
        0x41 => {
            cursor.i32()?;
        }
        0x42 => {
            cursor.i64()?;
        }
        0x43 => {
            cursor.bytes(4)?;
        }
        0x44 => {
            cursor.bytes(8)?;
        }
        0xd0 => heap(cursor)?,
        0xe3 | 0xe5 => {
            cursor.u32()?;
            handlers(cursor)?;
        }
        0xe4 => {
            indices(cursor, 2)?;
            handlers(cursor)?;
        }
        0xfb => match cursor.u32()? {
            0
            | 1
            | 6
            | 7
            | 11..=14
            | 16
            | 32..=34
            | 0x80..=0x82
            | 0x86
            | 0x87
            | 0x8b..=0x8e
            | 0x92
            | 0x9b => indices(cursor, 1)?,
            2..=5 | 8..=10 | 17..=19 => indices(cursor, 2)?,
            20..=23 | 35 | 36 => heap(cursor)?,
            24 | 25 | 37 | 38 => {
                let flags = cursor.byte()?;
                if flags > 3 {
                    return Err(error(cursor, "invalid cast flags"));
                }
                cursor.u32()?;
                heap(cursor)?;
                heap(cursor)?;
            }
            15
            | 26..=30
            | 0x83..=0x85
            | 0x88..=0x8a
            | 0x90
            | 0x91
            | 0x93
            | 0x98..=0x9a
            | 0x9c
            | 0xa0..=0xa4
            | 0xb0..=0xb7 => {}
            _ => return Err(error(cursor, "unknown GC instruction")),
        },
        0xfc => match cursor.u32()? {
            0..=7 => {}
            8 | 10 | 12 | 14 => indices(cursor, 2)?,
            9 | 11 | 13 | 15..=17 => indices(cursor, 1)?,
            _ => return Err(error(cursor, "unknown miscellaneous instruction")),
        },
        0xfd => {
            simd(cursor)?;
        }
        0xfe => match cursor.u32()? {
            0..=2 | 0x10..=0x4e => memarg(cursor)?,
            3 => {
                if cursor.byte()? != 0 {
                    return Err(error(cursor, "invalid atomic fence immediate"));
                }
            }
            _ => return Err(error(cursor, "unknown atomic instruction")),
        },
        _ => return Err(error(cursor, "unknown instruction")),
    }
    Ok(Instruction {
        start,
        end: cursor.position(),
        opcode,
        index,
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn immediate_bytes_are_not_instructions() {
        let bytes = [
            0x44, 0x20, 0x23, 0x0d, 0x03, 0x0b, 0x24, 0x21, 0x22, 0x20, 0x80, 1,
        ];
        let mut cursor = Cursor::new(&bytes);
        assert_eq!(next(&mut cursor).unwrap().end, 9);
        let op = next(&mut cursor).unwrap();
        assert_eq!(op.index, Some(128));
        assert_eq!(op.start, 9);
        assert_eq!(op.end, 12);
    }
    #[test]
    fn gc_simd_multimemory_and_catches_preserve_boundaries() {
        let bytes = [
            0xfb, 2, 0x80, 1, 3, 0xfd, 0x54, 0x40, 2, 0, 7, 0x1f, 0x40, 2, 0, 1, 0, 2, 0, 0xd0,
            0x62, 0x80, 1, 0x0b,
        ];
        let mut cursor = Cursor::new(&bytes);
        for end in [5, 11, 19, 23, 24] {
            assert_eq!(next(&mut cursor).unwrap().end, end);
        }
    }
    #[test]
    fn malformed_immediates_fail_instead_of_desynchronizing() {
        for bytes in [
            &[0x20, 0x80][..],
            &[0xfd, 12, 0],
            &[0xfb, 24, 4],
            &[0x0e, 0xff, 0xff, 0xff, 0xff, 15],
            &[0xfe, 3, 1],
            &[0xff],
        ] {
            assert!(next(&mut Cursor::new(bytes)).is_err(), "{bytes:?}");
        }
    }
    #[test]
    fn typed_blocks_and_continuation_handlers() {
        let bytes = [0x02, 0x63, 0x80, 1, 0xe3, 0, 2, 0, 1, 2, 1, 3, 0x0b];
        let mut cursor = Cursor::new(&bytes);
        assert_eq!(next(&mut cursor).unwrap().end, 4);
        assert_eq!(next(&mut cursor).unwrap().end, 12);
        assert_eq!(next(&mut cursor).unwrap().opcode, 0x0b);
    }
}
