//! Borrowed section framing shared by readers and binary inspection tools.
use crate::binary_cursor::Cursor;
pub use crate::binary_cursor::DecodeError;
use std::ops::Range;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BinaryKind {
    Module,
    Component,
}

#[derive(Debug)]
pub struct Section<'a> {
    pub id: u8,
    pub payload: &'a [u8],
    pub range: Range<usize>,
}

/// Reads section framing, not section typing or declaration order.
pub struct Sections<'a> {
    cursor: Cursor<'a>,
    kind: BinaryKind,
    failed: bool,
}

impl<'a> Sections<'a> {
    pub fn new(bytes: &'a [u8]) -> Result<Self, DecodeError> {
        let mut cursor = Cursor::new(bytes);
        let header = cursor.bytes(8)?;
        let kind = match header {
            b"\0asm\x01\0\0\0" => BinaryKind::Module,
            value if value == crate::component_format::HEADER => BinaryKind::Component,
            _ => {
                return Err(DecodeError {
                    offset: 0,
                    reason: "unsupported Wasm header",
                });
            }
        };
        Ok(Self {
            cursor,
            kind,
            failed: false,
        })
    }

    pub fn kind(&self) -> BinaryKind {
        self.kind
    }
}

impl<'a> Iterator for Sections<'a> {
    type Item = Result<Section<'a>, DecodeError>;

    fn next(&mut self) -> Option<Self::Item> {
        if self.failed || self.cursor.remaining() == 0 {
            return None;
        }
        let result = (|| {
            let id = self.cursor.byte()?;
            let length = self.cursor.u32()? as usize;
            let start = self.cursor.position();
            let payload = self.cursor.bytes(length)?;
            Ok(Section {
                id,
                payload,
                range: start..self.cursor.position(),
            })
        })();
        self.failed = result.is_err();
        Some(result)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn sections_borrow_payload_and_keep_absolute_offsets() {
        let bytes = b"\0asm\x01\0\0\0\0\x82\0xy\x01\0";
        let mut sections = Sections::new(bytes).unwrap();
        assert_eq!(sections.kind(), BinaryKind::Module);
        let first = sections.next().unwrap().unwrap();
        assert_eq!(first.id, 0);
        assert_eq!(first.range, 11..13);
        assert_eq!(first.payload.as_ptr(), bytes[11..].as_ptr());
        assert_eq!(sections.next().unwrap().unwrap().range, 15..15);
        assert!(sections.next().is_none());
        assert_eq!(
            Sections::new(crate::component_format::HEADER)
                .unwrap()
                .kind(),
            BinaryKind::Component
        );
    }

    #[test]
    fn truncated_section_stops_iteration_without_recovery() {
        for suffix in [&[1][..], &[1, 0x80], &[1, 2, 0]] {
            let mut bytes = b"\0asm\x01\0\0\0".to_vec();
            bytes.extend_from_slice(suffix);
            let mut sections = Sections::new(&bytes).unwrap();
            assert!(sections.next().unwrap().is_err());
            assert!(sections.next().is_none());
        }
        assert!(Sections::new(b"\0asm\x02\0\0\0").is_err());
    }
}
