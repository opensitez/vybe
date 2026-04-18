use std::path::Path;

/// Read a text file with BOM-aware encoding detection.
pub fn read_text_file(path: impl AsRef<Path>) -> std::io::Result<String> {
    let bytes = std::fs::read(path)?;
    Ok(decode_with_bom(&bytes))
}

pub fn decode_with_bom(bytes: &[u8]) -> String {
    if bytes.len() >= 4 && bytes[0] == 0xFF && bytes[1] == 0xFE && bytes[2] == 0x00 && bytes[3] == 0x00 {
        return decode_utf32_le(&bytes[4..]);
    }
    if bytes.len() >= 4 && bytes[0] == 0x00 && bytes[1] == 0x00 && bytes[2] == 0xFE && bytes[3] == 0xFF {
        return decode_utf32_be(&bytes[4..]);
    }
    if bytes.len() >= 2 && bytes[0] == 0xFF && bytes[1] == 0xFE {
        return decode_utf16_le(&bytes[2..]);
    }
    if bytes.len() >= 2 && bytes[0] == 0xFE && bytes[1] == 0xFF {
        return decode_utf16_be(&bytes[2..]);
    }
    if bytes.len() >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF {
        return String::from_utf8_lossy(&bytes[3..]).into_owned();
    }
    String::from_utf8_lossy(bytes).into_owned()
}

fn decode_utf16_le(bytes: &[u8]) -> String {
    let u16_iter = bytes.chunks_exact(2).map(|c| u16::from_le_bytes([c[0], c[1]]));
    char::decode_utf16(u16_iter).map(|r| r.unwrap_or('\u{FFFD}')).collect()
}

fn decode_utf16_be(bytes: &[u8]) -> String {
    let u16_iter = bytes.chunks_exact(2).map(|c| u16::from_be_bytes([c[0], c[1]]));
    char::decode_utf16(u16_iter).map(|r| r.unwrap_or('\u{FFFD}')).collect()
}

fn decode_utf32_le(bytes: &[u8]) -> String {
    bytes.chunks_exact(4)
        .filter_map(|c| char::from_u32(u32::from_le_bytes([c[0], c[1], c[2], c[3]])))
        .collect()
}

fn decode_utf32_be(bytes: &[u8]) -> String {
    bytes.chunks_exact(4)
        .filter_map(|c| char::from_u32(u32::from_be_bytes([c[0], c[1], c[2], c[3]])))
        .collect()
}
