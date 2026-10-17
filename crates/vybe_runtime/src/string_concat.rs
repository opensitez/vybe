//! Allocate shared UTF-8 concatenations directly in their final Arc storage.
use std::sync::Arc;

#[inline]
fn concat_parts_arc(parts: &[&str], len: usize) -> Arc<str> {
    let allocation = Arc::<[u8]>::new_uninit_slice(len);
    let raw = Arc::into_raw(allocation);
    let mut offset = 0;
    // Each source is valid UTF-8. The checked total length covers all copies,
    // and the new allocation cannot overlap any borrowed source. Every byte
    // is initialized exactly once, including the zero-length case.
    unsafe {
        // This freshly allocated Arc has never been cloned or published. Its
        // sole ownership is held by raw until the final from_raw below, so no
        // strong/weak-count uniqueness check is needed before initialization.
        let destination = raw as *mut u8;
        for part in parts {
            std::ptr::copy_nonoverlapping(part.as_ptr(), destination.add(offset), part.len());
            offset += part.len();
        }
        // [MaybeUninit<u8>] and str have identical allocation layout and length
        // metadata. The copies have initialized all bytes before publishing.
        // Concatenation of complete valid UTF-8 inputs is valid UTF-8. Transfer
        // the single Arc ownership without changing its allocation/refcount.
        Arc::from_raw(raw as *const str)
    }
}

#[inline]
pub fn concat2_arc(left: &str, right: &str) -> Arc<str> {
    if left.is_empty() {
        return Arc::from(right);
    }
    if right.is_empty() {
        return Arc::from(left);
    }
    let len = left.len().checked_add(right.len()).expect("string length overflow");
    concat_parts_arc(&[left, right], len)
}

#[inline]
pub fn concat3_arc(left: &str, middle: &str, right: &str) -> Arc<str> {
    if left.is_empty() {
        return concat2_arc(middle, right);
    }
    if middle.is_empty() {
        return concat2_arc(left, right);
    }
    if right.is_empty() {
        return concat2_arc(left, middle);
    }
    let len = left.len().checked_add(middle.len())
        .and_then(|len| len.checked_add(right.len()))
        .expect("string length overflow");
    concat_parts_arc(&[left, middle, right], len)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn final_arc_storage_preserves_unicode_nuls_and_empty_parts() {
        let parts = ["", "A", "\0", "Ωé中", "😀𐐀", "A\0Ω😀"];
        for left in parts {
            for right in parts {
                let expected = format!("{left}{right}");
                let output = concat2_arc(left, right);
                assert_eq!(output.as_bytes(), expected.as_bytes());
                assert_eq!(Arc::strong_count(&output), 1);
                for middle in parts {
                    let expected = format!("{left}{middle}{right}");
                    let output = concat3_arc(left, middle, right);
                    assert_eq!(output.as_bytes(), expected.as_bytes());
                    assert_eq!(Arc::strong_count(&output), 1);
                }
            }
        }
    }

    #[test]
    fn copied_parts_preserve_unaligned_sources_and_bulk_boundaries() {
        let source: String = (0..8192)
            .map(|index| char::from(b' ' + (index % 95) as u8))
            .collect();
        for offset in 0..16 {
            for left_len in [0, 1, 7, 8, 15, 16, 17, 63, 64, 65, 4095, 4096] {
                for right_len in [0, 1, 15, 16, 17, 63, 64, 65] {
                    let left = &source[offset..offset + left_len];
                    let right = &source[31..31 + right_len];
                    let output = concat3_arc(left, "Ω\0😀", right);
                    let expected = format!("{left}Ω\0😀{right}");
                    assert_eq!(output.as_bytes(), expected.as_bytes());
                }
            }
        }
    }

    #[test]
    fn output_owns_its_bytes_and_uses_normal_arc_lifetime() {
        let output = {
            let source: Arc<str> = Arc::from("A\0Ω😀");
            concat3_arc(&source, &source, &source)
        };
        assert_eq!(&*output, "A\0Ω😀A\0Ω😀A\0Ω😀");
        let weak = Arc::downgrade(&output);
        let alias = Arc::clone(&output);
        assert_eq!(Arc::strong_count(&alias), 2);
        drop(output);
        assert_eq!(Arc::strong_count(&alias), 1);
        assert_eq!(weak.upgrade().as_deref(), Some(&*alias));
        drop(alias);
        assert!(weak.upgrade().is_none());
    }
}

