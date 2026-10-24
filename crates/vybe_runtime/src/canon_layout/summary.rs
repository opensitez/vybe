//! Checked memory64 layout summaries for canonical value-definition validation.
//! Inputs can be borrowed or cached summaries; no type tree is retained.

const SIZE_ERROR: &str = "exceeds maximum byte size";

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct CanonicalLayout {
    size: u64,
    align: u64,
}

impl CanonicalLayout {
    pub const BYTE: Self = Self { size: 1, align: 1 };
    pub const WORD: Self = Self { size: 2, align: 2 };
    pub const DWORD: Self = Self { size: 4, align: 4 };
    pub const QWORD: Self = Self { size: 8, align: 8 };
    /// Memory64 string/list pointer and length pair.
    pub const POINTER_PAIR: Self = Self { size: 16, align: 8 };

    pub const fn size(self) -> u64 {
        self.size
    }
    pub const fn alignment(self) -> u64 {
        self.align
    }

    #[inline]
    pub fn validate_value_size(self) -> Result<(), &'static str> {
        if self.size < 1 << 28 {
            Ok(())
        } else {
            Err(SIZE_ERROR)
        }
    }

    #[inline]
    pub fn flags(count: usize) -> Result<Self, &'static str> {
        match count {
            1..=8 => Ok(Self::BYTE),
            9..=16 => Ok(Self::WORD),
            17..=32 => Ok(Self::DWORD),
            _ => Err("component flags require 1 to 32 labels"),
        }
    }

    #[inline]
    pub fn fixed_list(self, count: u32) -> Result<Self, &'static str> {
        if count == 0 {
            return Err("fixed-length list must not be empty");
        }
        Ok(Self {
            size: self.size.checked_mul(u64::from(count)).ok_or(SIZE_ERROR)?,
            align: self.align,
        })
    }

    #[inline]
    pub fn record(
        fields: impl IntoIterator<Item = Result<Self, &'static str>>,
    ) -> Result<Self, &'static str> {
        let mut fields = fields.into_iter();
        let first = fields.next().ok_or("component record requires a field")??;
        let mut size = first.size;
        let mut align = first.align;
        for field in fields {
            let field = field?;
            size = checked_align(size, field.align)?
                .checked_add(field.size)
                .ok_or(SIZE_ERROR)?;
            align = align.max(field.align);
        }
        Ok(Self {
            size: checked_align(size, align)?,
            align,
        })
    }

    #[inline]
    pub fn variant(
        cases: impl ExactSizeIterator<Item = Result<Option<Self>, &'static str>>,
    ) -> Result<Self, &'static str> {
        let tag = match cases.len() {
            0 => return Err("component variant requires a case"),
            1..=256 => 1,
            257..=65_536 => 2,
            n if n <= u32::MAX as usize => 4,
            _ => return Err("component variant has too many cases"),
        };
        let mut payload_align = 1;
        let mut payload_size = 0;
        for case in cases {
            if let Some(case) = case? {
                payload_align = payload_align.max(case.align);
                payload_size = payload_size.max(case.size);
            }
        }
        let align = tag.max(payload_align);
        let end = checked_align(tag, payload_align)?
            .checked_add(payload_size)
            .ok_or(SIZE_ERROR)?;
        Ok(Self {
            size: checked_align(end, align)?,
            align,
        })
    }
}

fn checked_align(size: u64, align: u64) -> Result<u64, &'static str> {
    size.checked_add(align - 1)
        .map(|n| n / align * align)
        .ok_or(SIZE_ERROR)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn variant_discriminant_boundaries() {
        for (count, size) in [(1, 1), (256, 1), (257, 2), (65536, 2), (65537, 4)] {
            let layout = CanonicalLayout::variant((0..count).map(|_| Ok(None))).unwrap();
            assert_eq!((layout.size(), layout.alignment()), (size, size));
        }
    }

    #[test]
    fn checked_padding_and_limits() {
        let record =
            CanonicalLayout::record([Ok(CanonicalLayout::BYTE), Ok(CanonicalLayout::QWORD)])
                .unwrap();
        assert_eq!((record.size(), record.alignment()), (16, 8));
        assert!(
            CanonicalLayout::BYTE
                .fixed_list((1 << 28) - 1)
                .unwrap()
                .validate_value_size()
                .is_ok()
        );
        assert_eq!(
            CanonicalLayout::BYTE
                .fixed_list(1 << 28)
                .unwrap()
                .validate_value_size(),
            Err(SIZE_ERROR)
        );
        let huge = CanonicalLayout::QWORD
            .fixed_list(u32::MAX)
            .unwrap()
            .fixed_list(u32::MAX);
        assert_eq!(huge, Err(SIZE_ERROR));
        assert!(CanonicalLayout::record([]).is_err());
        assert!(CanonicalLayout::variant(std::iter::empty()).is_err());
    }

    #[test]
    fn record_first_field_retains_alignment_and_errors() {
        for field in [CanonicalLayout::BYTE, CanonicalLayout::QWORD, CanonicalLayout::POINTER_PAIR] {
            assert_eq!(CanonicalLayout::record([Ok(field)]), Ok(field));
        }
        assert_eq!(CanonicalLayout::record([Err("first"), Ok(CanonicalLayout::BYTE)]), Err("first"));
        assert_eq!(CanonicalLayout::record([Ok(CanonicalLayout::BYTE), Err("second")]), Err("second"));
    }
}
