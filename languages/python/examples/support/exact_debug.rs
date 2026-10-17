//! Exact Debug comparison without allocating a complete formatted AST String.
use std::fmt::{self, Debug};
use std::io::{self, BufReader, BufWriter, Read, Seek, SeekFrom, Write};
use std::path::PathBuf;
use std::sync::atomic::{AtomicU64, Ordering};

pub struct Snapshot {
    path: PathBuf,
}

#[derive(Debug, PartialEq, Eq)]
pub struct Difference {
    pub byte_offset: u64,
    pub expected: Option<u8>,
    pub actual: Option<u8>,
}

impl Snapshot {
    pub fn new(value: &impl Debug) -> io::Result<Self> {
        static NEXT: AtomicU64 = AtomicU64::new(0);
        let path = std::env::temp_dir().join(format!(
            "vybe-python-ast-{}-{}.debug",
            std::process::id(),
            NEXT.fetch_add(1, Ordering::Relaxed)
        ));
        let file = std::fs::OpenOptions::new()
            .write(true)
            .create_new(true)
            .open(&path)?;
        let snapshot = Self { path };
        let mut output = BufWriter::new(file);
        write!(output, "{value:?}")?;
        output.flush()?;
        Ok(snapshot)
    }

    pub fn compare(&self, value: &impl Debug) -> io::Result<Option<Difference>> {
        let mut input = BufReader::new(std::fs::File::open(&self.path)?);
        input.seek(SeekFrom::Start(0))?;
        let mut output = Comparison {
            input: &mut input,
            offset: 0,
            difference: None,
            error: None,
        };
        if fmt::write(&mut output, format_args!("{value:?}")).is_err() {
            if let Some(error) = output.error {
                return Err(error);
            }
            if output.difference.is_none() {
                return Err(io::Error::other("Debug formatting failed"));
            }
        }
        if output.difference.is_some() {
            return Ok(output.difference);
        }
        let mut tail = [0];
        if output.input.read(&mut tail)? != 0 {
            return Ok(Some(Difference {
                byte_offset: output.offset,
                expected: Some(tail[0]),
                actual: None,
            }));
        }
        Ok(None)
    }
}

impl Drop for Snapshot {
    fn drop(&mut self) {
        let _ = std::fs::remove_file(&self.path);
    }
}

struct Comparison<'a> {
    input: &'a mut BufReader<std::fs::File>,
    offset: u64,
    difference: Option<Difference>,
    error: Option<io::Error>,
}
impl fmt::Write for Comparison<'_> {
    fn write_str(&mut self, text: &str) -> fmt::Result {
        let mut bytes = text.as_bytes();
        let mut expected = [0; 4096];
        while !bytes.is_empty() {
            let wanted = bytes.len().min(expected.len());
            let read = match self.input.read(&mut expected[..wanted]) {
                Ok(read) => read,
                Err(error) => {
                    self.error = Some(error);
                    return Err(fmt::Error);
                }
            };
            let mismatch = expected[..read].iter().zip(bytes).position(|(a, b)| a != b);
            if let Some(index) = mismatch {
                self.difference = Some(Difference {
                    byte_offset: self.offset + index as u64,
                    expected: Some(expected[index]),
                    actual: Some(bytes[index]),
                });
                return Err(fmt::Error);
            }
            self.offset += read as u64;
            if read == 0 {
                self.difference = Some(Difference {
                    byte_offset: self.offset,
                    expected: None,
                    actual: Some(bytes[0]),
                });
                return Err(fmt::Error);
            }
            bytes = &bytes[read..];
        }
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    struct Text<'a>(&'a str, usize);
    impl Debug for Text<'_> {
        fn fmt(&self, output: &mut fmt::Formatter<'_>) -> fmt::Result {
            for chunk in self.0.as_bytes().chunks(self.1) {
                output.write_str(std::str::from_utf8(chunk).unwrap())?;
            }
            Ok(())
        }
    }
    #[test]
    fn chunk_boundaries_do_not_change_exact_equality() {
        let text = "0123456789".repeat(1500);
        let snapshot = Snapshot::new(&Text(&text, 7000)).unwrap();
        assert_eq!(snapshot.compare(&Text(&text, 13)).unwrap(), None);
        let path = snapshot.path.clone();
        drop(snapshot);
        assert!(!path.exists());
    }
    #[test]
    fn finds_first_difference_and_both_length_mismatches() {
        let snapshot = Snapshot::new(&Text("abc", 2)).unwrap();
        for (actual, expected) in [
            (
                "abd",
                Difference {
                    byte_offset: 2,
                    expected: Some(b'c'),
                    actual: Some(b'd'),
                },
            ),
            (
                "ab",
                Difference {
                    byte_offset: 2,
                    expected: Some(b'c'),
                    actual: None,
                },
            ),
            (
                "abcd",
                Difference {
                    byte_offset: 3,
                    expected: None,
                    actual: Some(b'd'),
                },
            ),
        ] {
            assert_eq!(snapshot.compare(&Text(actual, 1)).unwrap(), Some(expected));
        }
    }
}
