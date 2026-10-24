//! Bounded pipe reads and diagnostic retention, independent of line length.

use std::io::{self, BufRead};

pub const OUTPUT_CHUNK_BYTES: usize = 8192;
pub const MAX_RETAINED_OUTPUT: usize = 65536;

pub struct OutputEvent {
    /// Raw bytes, including the newline when `end_of_line` is true.
    pub bytes: Vec<u8>,
    pub end_of_line: bool,
}

pub fn read_event<R: BufRead>(reader: &mut R) -> io::Result<Option<OutputEvent>> {
    loop {
        let available = match reader.fill_buf() {
            Err(error) if error.kind() == io::ErrorKind::Interrupted => continue,
            result => result?,
        };
        if available.is_empty() {
            return Ok(None);
        }
        let available = &available[..available.len().min(OUTPUT_CHUNK_BYTES)];
        let newline = available.iter().position(|byte| *byte == b'\n');
        let len = newline.map_or(available.len(), |index| index + 1);
        let event = OutputEvent {
            bytes: available[..len].to_vec(),
            end_of_line: newline.is_some(),
        };
        reader.consume(len);
        return Ok(Some(event));
    }
}

#[derive(Default)]
pub struct BoundedOutput {
    bytes: Vec<u8>,
}

impl BoundedOutput {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn push(&mut self, bytes: &[u8]) {
        let remaining = MAX_RETAINED_OUTPUT.saturating_sub(self.bytes.len());
        self.bytes
            .extend_from_slice(&bytes[..bytes.len().min(remaining)]);
    }

    pub fn text(&self) -> String {
        // Decode only after retention so split UTF-8 sequences remain intact.
        String::from_utf8_lossy(&self.bytes).into_owned()
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::{BufReader, Cursor};

    #[test]
    fn huge_unterminated_line_is_fragmented_and_fully_drained() {
        let input = vec![b'x'; MAX_RETAINED_OUTPUT * 8 + 3];
        let mut reader = BufReader::with_capacity(OUTPUT_CHUNK_BYTES * 4, Cursor::new(&input));
        let mut retained = BoundedOutput::new();
        let mut total = 0;
        while let Some(event) = read_event(&mut reader).unwrap() {
            assert!(event.bytes.len() <= OUTPUT_CHUNK_BYTES);
            assert!(!event.end_of_line);
            total += event.bytes.len();
            retained.push(&event.bytes);
        }
        assert_eq!(total, input.len());
        assert_eq!(retained.bytes.len(), MAX_RETAINED_OUTPUT);
    }

    #[test]
    fn newline_binary_and_split_utf8_are_preserved() {
        let input = b"a\r\n\xff\n\xe2\x82\xac";
        let mut reader = BufReader::with_capacity(1, Cursor::new(input));
        let mut retained = BoundedOutput::new();
        let mut newlines = 0;
        while let Some(event) = read_event(&mut reader).unwrap() {
            newlines += usize::from(event.end_of_line);
            retained.push(&event.bytes);
        }
        assert_eq!(newlines, 2);
        assert_eq!(retained.bytes, input);
        assert!(retained.text().ends_with('\u{20ac}'));
    }
}
