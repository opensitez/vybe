use std::path::Path;

/// Read a text file with proper encoding detection based on BOM (Byte Order Mark).
///
/// Supports:
/// - UTF-8 with BOM (EF BB BF) → decoded as UTF-8, BOM stripped
/// - UTF-16 LE with BOM (FF FE) → decoded from UTF-16 LE, BOM stripped
/// - UTF-16 BE with BOM (FE FF) → decoded from UTF-16 BE, BOM stripped
/// - No BOM → assumed UTF-8 (most common for modern VB.NET files)
///
/// Visual Studio may save .vb files in any of these encodings.
/// This function ensures correct decoding regardless of encoding.
pub fn read_text_file(path: impl AsRef<Path>) -> std::io::Result<String> {
    let bytes = std::fs::read(path)?;
    Ok(decode_with_bom(&bytes))
}

/// Decode raw bytes to a String, detecting encoding from BOM.
/// Strips the BOM character from the result.
pub fn decode_with_bom(bytes: &[u8]) -> String {
    // UTF-32 LE BOM: FF FE 00 00 (check before UTF-16 LE since it shares the first 2 bytes)
    if bytes.len() >= 4 && bytes[0] == 0xFF && bytes[1] == 0xFE && bytes[2] == 0x00 && bytes[3] == 0x00 {
        return decode_utf32_le(&bytes[4..]);
    }

    // UTF-32 BE BOM: 00 00 FE FF
    if bytes.len() >= 4 && bytes[0] == 0x00 && bytes[1] == 0x00 && bytes[2] == 0xFE && bytes[3] == 0xFF {
        return decode_utf32_be(&bytes[4..]);
    }

    // UTF-16 LE BOM: FF FE
    if bytes.len() >= 2 && bytes[0] == 0xFF && bytes[1] == 0xFE {
        return decode_utf16_le(&bytes[2..]);
    }

    // UTF-16 BE BOM: FE FF
    if bytes.len() >= 2 && bytes[0] == 0xFE && bytes[1] == 0xFF {
        return decode_utf16_be(&bytes[2..]);
    }

    // UTF-8 BOM: EF BB BF
    if bytes.len() >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF {
        return String::from_utf8_lossy(&bytes[3..]).into_owned();
    }

    // No BOM — assume UTF-8
    String::from_utf8_lossy(bytes).into_owned()
}

fn decode_utf16_le(bytes: &[u8]) -> String {
    let u16_iter = bytes.chunks_exact(2).map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]));
    char::decode_utf16(u16_iter)
        .map(|r| r.unwrap_or('\u{FFFD}'))
        .collect()
}

fn decode_utf16_be(bytes: &[u8]) -> String {
    let u16_iter = bytes.chunks_exact(2).map(|chunk| u16::from_be_bytes([chunk[0], chunk[1]]));
    char::decode_utf16(u16_iter)
        .map(|r| r.unwrap_or('\u{FFFD}'))
        .collect()
}

fn decode_utf32_le(bytes: &[u8]) -> String {
    bytes
        .chunks_exact(4)
        .filter_map(|chunk| {
            let code = u32::from_le_bytes([chunk[0], chunk[1], chunk[2], chunk[3]]);
            char::from_u32(code)
        })
        .collect()
}

fn decode_utf32_be(bytes: &[u8]) -> String {
    bytes
        .chunks_exact(4)
        .filter_map(|chunk| {
            let code = u32::from_be_bytes([chunk[0], chunk[1], chunk[2], chunk[3]]);
            char::from_u32(code)
        })
        .collect()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_utf8_no_bom() {
        let bytes = b"Hello World";
        assert_eq!(decode_with_bom(bytes), "Hello World");
    }

    #[test]
    fn test_utf8_with_bom() {
        let bytes = b"\xEF\xBB\xBFHello World";
        assert_eq!(decode_with_bom(bytes), "Hello World");
    }

    #[test]
    fn test_utf16_le_with_bom() {
        // "Hi" in UTF-16 LE with BOM
        let bytes: Vec<u8> = vec![
            0xFF, 0xFE, // BOM
            b'H', 0x00, b'i', 0x00,
        ];
        assert_eq!(decode_with_bom(&bytes), "Hi");
    }

    #[test]
    fn test_utf16_be_with_bom() {
        // "Hi" in UTF-16 BE with BOM
        let bytes: Vec<u8> = vec![
            0xFE, 0xFF, // BOM
            0x00, b'H', 0x00, b'i',
        ];
        assert_eq!(decode_with_bom(&bytes), "Hi");
    }

    #[test]
    fn test_utf8_bom_stripped_no_residue() {
        let bytes = b"\xEF\xBB\xBFImports System";
        let result = decode_with_bom(bytes);
        assert!(!result.starts_with('\u{feff}'));
        assert_eq!(result, "Imports System");
    }
}
