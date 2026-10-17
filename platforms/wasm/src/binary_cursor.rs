//! Checked, borrowing binary primitives for the in-house Wasm decoder.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeError {
    pub offset: usize,
    pub reason: &'static str,
}
impl std::fmt::Display for DecodeError {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(formatter, "{} at byte {}", self.reason, self.offset)
    }
}
impl std::error::Error for DecodeError {}

pub struct Cursor<'a> {
    bytes: &'a [u8],
    position: usize,
}

impl<'a> Cursor<'a> {
    pub fn new(bytes: &'a [u8]) -> Self {
        Self { bytes, position: 0 }
    }
    pub fn position(&self) -> usize {
        self.position
    }
    pub fn remaining(&self) -> usize {
        self.bytes.len() - self.position
    }
    pub fn peek(&self) -> Option<u8> {
        self.bytes.get(self.position).copied()
    }
    pub fn byte(&mut self) -> Result<u8, DecodeError> {
        let value = self.bytes.get(self.position).copied().ok_or(DecodeError {
            offset: self.position,
            reason: "unexpected end of binary",
        })?;
        self.position += 1;
        Ok(value)
    }
    pub fn bytes(&mut self, count: usize) -> Result<&'a [u8], DecodeError> {
        let end = self
            .position
            .checked_add(count)
            .filter(|end| *end <= self.bytes.len())
            .ok_or(DecodeError {
                offset: self.position,
                reason: "unexpected end of binary",
            })?;
        let result = &self.bytes[self.position..end];
        self.position = end;
        Ok(result)
    }
    pub fn unsigned(&mut self, bits: u32) -> Result<u64, DecodeError> {
        assert!((1..=64).contains(&bits));
        let max_bytes = bits.div_ceil(7);
        let mut value = 0u64;
        for index in 0..max_bytes {
            let byte = self.byte()?;
            let payload = byte & 0x7f;
            let shift = index * 7;
            let available = (bits - shift).min(7);
            if available < 7 && u32::from(payload) >= (1u32 << available) {
                return Err(DecodeError {
                    offset: self.position - 1,
                    reason: "integer too large",
                });
            }
            value |= u64::from(payload) << shift;
            if byte & 0x80 == 0 {
                return Ok(value);
            }
        }
        Err(DecodeError {
            offset: self.position - 1,
            reason: "integer representation too long",
        })
    }
    pub fn signed(&mut self, bits: u32) -> Result<i64, DecodeError> {
        assert!((1..=64).contains(&bits));
        let max_bytes = bits.div_ceil(7);
        let mut value = 0u64;
        for index in 0..max_bytes {
            let byte = self.byte()?;
            let payload = byte & 0x7f;
            let shift = index * 7;
            let available = (bits - shift).min(7);
            if available < 7 {
                let sign = 1u8 << (available - 1);
                let unused = 0x7f & !((1u8 << available) - 1);
                let expected = if payload & sign == 0 { 0 } else { unused };
                if payload & unused != expected {
                    return Err(DecodeError {
                        offset: self.position - 1,
                        reason: "integer too large",
                    });
                }
            }
            value |= u64::from(payload) << shift;
            if byte & 0x80 == 0 {
                let consumed_bits = shift + 7;
                if consumed_bits < 64 && payload & 0x40 != 0 {
                    value |= u64::MAX << consumed_bits;
                }
                return Ok(value as i64);
            }
        }
        Err(DecodeError {
            offset: self.position - 1,
            reason: "integer representation too long",
        })
    }
    pub fn u32(&mut self) -> Result<u32, DecodeError> {
        self.unsigned(32).map(|v| v as u32)
    }
    pub fn u64(&mut self) -> Result<u64, DecodeError> {
        self.unsigned(64)
    }
    pub fn i32(&mut self) -> Result<i32, DecodeError> {
        self.signed(32).map(|v| v as i32)
    }
    pub fn i64(&mut self) -> Result<i64, DecodeError> {
        self.signed(64)
    }
    pub fn s33(&mut self) -> Result<i64, DecodeError> {
        self.signed(33)
    }
    pub fn name(&mut self) -> Result<&'a str, DecodeError> {
        let count = self.u32()? as usize;
        let start = self.position;
        std::str::from_utf8(self.bytes(count)?).map_err(|_| DecodeError {
            offset: start,
            reason: "invalid UTF-8 name",
        })
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn unsigned_boundaries_and_nonminimal_encodings() {
        for (bytes, expected) in [
            (&[0][..], 0),
            (&[0x80, 0], 0),
            (&[0xff, 0xff, 0xff, 0xff, 0x0f], u32::MAX),
        ] {
            assert_eq!(Cursor::new(bytes).u32().unwrap(), expected);
        }
        assert_eq!(
            Cursor::new(&[0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 1])
                .u64()
                .unwrap(),
            u64::MAX
        );
    }
    #[test]
    fn signed_boundaries_and_valid_padding() {
        for (bytes, expected) in [
            (&[0x7f][..], -1),
            (&[0xff, 0x7f], -1),
            (&[0x80, 0x80, 0x80, 0x80, 0x78], i32::MIN),
            (&[0xff, 0xff, 0xff, 0xff, 7], i32::MAX),
        ] {
            assert_eq!(Cursor::new(bytes).i32().unwrap(), expected);
        }
        assert_eq!(
            Cursor::new(&[0xff, 0xff, 0xff, 0xff, 0x0f]).s33().unwrap(),
            u32::MAX as i64
        );
        assert_eq!(
            Cursor::new(&[0x80, 0x80, 0x80, 0x80, 0x70]).s33().unwrap(),
            -(1i64 << 32)
        );
        assert_eq!(
            Cursor::new(&[0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0])
                .i64()
                .unwrap(),
            i64::MAX
        );
        assert_eq!(
            Cursor::new(&[0x80, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80, 0x7f])
                .i64()
                .unwrap(),
            i64::MIN
        );
    }
    #[test]
    fn malformed_terminal_bits_and_truncation() {
        for bytes in [
            &[][..],
            &[0x80],
            &[0x80, 0x80, 0x80, 0x80, 0x80],
            &[0xff, 0xff, 0xff, 0xff, 0x10],
        ] {
            assert!(Cursor::new(bytes).u32().is_err());
        }
        for bytes in [
            &[0xff, 0xff, 0xff, 0xff, 0x0f][..],
            &[0x80, 0x80, 0x80, 0x80, 0x70],
        ] {
            assert!(Cursor::new(bytes).i32().is_err());
        }
        assert!(
            Cursor::new(&[0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 2])
                .u64()
                .is_err()
        );
        assert!(
            Cursor::new(&[0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 1])
                .i64()
                .is_err()
        );
    }
    #[test]
    fn borrowed_slices_names_and_failure_position() {
        let input = [3, b'a', 0, b'b', 0x41];
        let mut cursor = Cursor::new(&input);
        let name = cursor.name().unwrap();
        assert_eq!(name.as_bytes(), b"a\0b");
        assert_eq!(name.as_ptr(), input[1..].as_ptr());
        assert_eq!(cursor.position(), 4);
        assert_eq!(cursor.remaining(), 1);
        assert!(cursor.bytes(usize::MAX).is_err());
        assert_eq!(cursor.position(), 4);
        assert_eq!(cursor.byte().unwrap(), 0x41);
        assert!(cursor.byte().is_err());
        assert!(Cursor::new(&[1, 0xff]).name().is_err());
        assert!(Cursor::new(&[2, b'a']).name().is_err());
    }
}
