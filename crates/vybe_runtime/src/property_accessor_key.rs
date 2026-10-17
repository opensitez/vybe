//! Temporary property lookup keys, without heap allocation for common names.
//! This changes neither stored property names nor the guest object model.

pub(super) enum AccessorKey {
    Inline { bytes: [u8; 64], len: usize },
    Owned(String),
}

impl AccessorKey {
    #[inline]
    pub(super) fn getter(name: &str) -> Self {
        const PREFIX: &str = "__get_";
        if name.len() <= 64 - PREFIX.len() {
            let mut bytes = [0; 64];
            bytes[..PREFIX.len()].copy_from_slice(PREFIX.as_bytes());
            let len = PREFIX.len() + name.len();
            bytes[PREFIX.len()..len].copy_from_slice(name.as_bytes());
            Self::Inline { bytes, len }
        } else {
            let mut key = String::with_capacity(PREFIX.len() + name.len());
            key.push_str(PREFIX);
            key.push_str(name);
            Self::Owned(key)
        }
    }

    #[inline]
    pub(super) fn as_str(&self) -> &str {
        match self {
            Self::Inline { bytes, len } => std::str::from_utf8(&bytes[..*len])
                .expect("getter prefix and name preserve UTF-8"),
            Self::Owned(key) => key,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::AccessorKey;

    #[test]
    fn keys_preserve_exact_bytes_and_inline_boundary() {
        for len in 0..180 {
            for unit in ["a", "\0", "\u{e9}", "\u{4e2d}", "\u{1f600}"] {
                let name = unit.repeat(len);
                let key = AccessorKey::getter(&name);
                assert_eq!(key.as_str(), format!("__get_{name}"));
                assert_eq!(matches!(key, AccessorKey::Inline { .. }), name.len() <= 58);
            }
        }
    }

    #[test]
    fn temporary_keys_do_not_change_lookup_order_or_names() {
        let mut properties = std::collections::HashMap::new();
        properties.insert("pixels".to_owned(), 1);
        properties.insert("__get_pixels".to_owned(), 2);
        properties.insert("__get_Pixels".to_owned(), 3);
        for (name, expected) in [("pixels", Some(&2)), ("Pixels", Some(&3)), ("PIXELS", None)] {
            let key = AccessorKey::getter(name);
            assert_eq!(properties.get(key.as_str()), expected);
        }
        assert_eq!(properties.get("pixels"), Some(&1));
    }
}
