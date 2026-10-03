//! Indexed source coordinates for diagnostics, AST adapters and editor edits.
//! Sparse scalar-boundary checkpoints bound scans on very long Unicode lines.
use std::{ops::Range, sync::Arc};

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Encoding {
    Utf8,
    Utf16,
    Scalar,
}
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Position {
    pub line: u32,
    pub character: u32,
}
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PositionError {
    OutOfBounds,
    InsideScalar,
    InsideLineBreak,
    Overflow,
}
#[derive(Debug)]
struct Line {
    start: usize,
    end: usize,
    checkpoints: Range<usize>,
}
#[derive(Debug, Clone, Copy)]
struct Checkpoint {
    byte: usize,
    utf16: usize,
    scalars: usize,
}
#[derive(Debug)]
pub struct Index<'i> {
    source: &'i str,
    data: Arc<Coordinates>,
}
#[derive(Debug)]
struct Coordinates {
    lines: Vec<Line>,
    checkpoints: Vec<Checkpoint>,
}
#[derive(Debug, Clone)]
pub struct OwnedIndex {
    source: Arc<str>,
    data: Arc<Coordinates>,
}
impl OwnedIndex {
    pub fn new(source: Arc<str>) -> Self {
        let data = Index::new(&source).data;
        Self { source, data }
    }
    pub fn source(&self) -> &str {
        &self.source
    }
    pub fn index(&self) -> Index<'_> {
        Index {
            source: &self.source,
            data: self.data.clone(),
        }
    }
}

impl<'i> Index<'i> {
    pub fn new(source: &'i str) -> Self {
        let mut lines = Vec::new();
        let mut checkpoints = vec![Checkpoint {
            byte: 0,
            utf16: 0,
            scalars: 0,
        }];
        let mut start = 0;
        let mut checkpoint_start = 0;
        let mut byte = 0;
        let mut utf16 = 0;
        let mut scalars = 0;
        while byte < source.len() {
            let ch = source[byte..].chars().next().unwrap();
            if matches!(ch, '\r' | '\n') {
                lines.push(Line {
                    start,
                    end: byte,
                    checkpoints: checkpoint_start..checkpoints.len(),
                });
                byte += if ch == '\r' && source.as_bytes().get(byte + 1) == Some(&b'\n') {
                    2
                } else {
                    1
                };
                start = byte;
                utf16 = 0;
                scalars = 0;
                checkpoint_start = checkpoints.len();
                checkpoints.push(Checkpoint {
                    byte,
                    utf16,
                    scalars,
                });
            } else {
                byte += ch.len_utf8();
                utf16 += ch.len_utf16();
                scalars += 1;
                if byte - checkpoints.last().unwrap().byte >= 64 {
                    checkpoints.push(Checkpoint {
                        byte,
                        utf16,
                        scalars,
                    });
                }
            }
        }
        lines.push(Line {
            start,
            end: byte,
            checkpoints: checkpoint_start..checkpoints.len(),
        });
        Self {
            source,
            data: Arc::new(Coordinates { lines, checkpoints }),
        }
    }
    pub fn source(&self) -> &'i str {
        self.source
    }
    pub fn line_count(&self) -> usize {
        self.data.lines.len()
    }
    /// Zero-based positions. CRLF's interior byte has no representable editor
    /// position and is rejected; the start of a line break is the line's end.
    pub fn position(&self, byte: usize, encoding: Encoding) -> Result<Position, PositionError> {
        if byte > self.source.len() {
            return Err(PositionError::OutOfBounds);
        }
        if !self.source.is_char_boundary(byte) {
            return Err(PositionError::InsideScalar);
        }
        let line_id = self.data.lines.partition_point(|line| line.start <= byte) - 1;
        let line = &self.data.lines[line_id];
        if byte > line.end {
            return Err(PositionError::InsideLineBreak);
        }
        let checkpoints = &self.data.checkpoints[line.checkpoints.clone()];
        let cp = checkpoints[checkpoints.partition_point(|cp| cp.byte <= byte) - 1];
        let suffix = &self.source[cp.byte..byte];
        let character = match encoding {
            Encoding::Utf8 => byte - line.start,
            Encoding::Utf16 => cp.utf16 + suffix.chars().map(char::len_utf16).sum::<usize>(),
            Encoding::Scalar => cp.scalars + suffix.chars().count(),
        };
        Ok(Position {
            line: line_id.try_into().map_err(|_| PositionError::Overflow)?,
            character: character.try_into().map_err(|_| PositionError::Overflow)?,
        })
    }
    pub fn offset(&self, position: Position, encoding: Encoding) -> Result<usize, PositionError> {
        let line = self
            .data
            .lines
            .get(position.line as usize)
            .ok_or(PositionError::OutOfBounds)?;
        let target = position.character as usize;
        if encoding == Encoding::Utf8 {
            let byte = line
                .start
                .checked_add(target)
                .ok_or(PositionError::Overflow)?;
            if byte > line.end {
                return Err(PositionError::OutOfBounds);
            }
            return if self.source.is_char_boundary(byte) {
                Ok(byte)
            } else {
                Err(PositionError::InsideScalar)
            };
        }
        let checkpoints = &self.data.checkpoints[line.checkpoints.clone()];
        let column = |cp: &Checkpoint| match encoding {
            Encoding::Utf16 => cp.utf16,
            Encoding::Scalar => cp.scalars,
            Encoding::Utf8 => unreachable!(),
        };
        let cp = checkpoints[checkpoints.partition_point(|cp| column(cp) <= target) - 1];
        let mut current = column(&cp);
        let mut byte = cp.byte;
        for ch in self.source[byte..line.end].chars() {
            if current == target {
                return Ok(byte);
            }
            current += if encoding == Encoding::Utf16 {
                ch.len_utf16()
            } else {
                1
            };
            byte += ch.len_utf8();
            if current > target {
                return Err(PositionError::InsideScalar);
            }
        }
        if current == target {
            Ok(byte)
        } else {
            Err(PositionError::OutOfBounds)
        }
    }
}
