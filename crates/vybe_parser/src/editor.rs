//! Revisioned editor foundation and full-reparse correctness baseline.
//! This does not claim incremental syntax reuse or error recovery. Reports own
//! source/index/captures, so an asynchronous client cannot mix revisions.
use std::sync::{
    Arc,
    atomic::{AtomicU64, Ordering},
};
static NEXT_DOCUMENT: AtomicU64 = AtomicU64::new(1);
use crate::{
    CaptureRule, CompiledGrammar, MatchOptions, ParseError, Span,
    source::{Encoding, OwnedIndex, Position, PositionError},
    tree::Node,
};

#[derive(Debug, Clone, Copy)]
pub struct TextRange {
    pub start: Position,
    pub end: Position,
}
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum EditError {
    StaleRevision,
    Position(PositionError),
    ReversedRange,
    RevisionOverflow,
}
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct NodeId {
    pub document: u64,
    pub revision: u64,
    pub index: usize,
}
#[derive(Debug, Clone, Copy)]
pub struct Capture {
    pub id: NodeId,
    pub rule: CaptureRule,
    pub span: Span,
    pub subtree_end: usize,
}

#[derive(Debug, Clone)]
pub struct Report {
    document: u64,
    revision: u64,
    grammar: Arc<CompiledGrammar>,
    source: OwnedIndex,
    nodes: Arc<[Node]>,
    consumed: Option<usize>,
    error: Option<ParseError>,
}
impl Report {
    pub fn revision(&self) -> u64 {
        self.revision
    }
    pub fn source(&self) -> &str {
        self.source.source()
    }
    pub fn index(&self) -> crate::source::Index<'_> {
        self.source.index()
    }
    pub fn consumed(&self) -> Option<usize> {
        self.consumed
    }
    pub fn error(&self) -> Option<&ParseError> {
        self.error.as_ref()
    }
    pub fn captures(&self) -> impl Iterator<Item = Capture> + '_ {
        self.nodes.iter().enumerate().map(|(index, node)| Capture {
            id: NodeId {
                document: self.document,
                revision: self.revision,
                index,
            },
            rule: node.rule,
            span: node.span,
            subtree_end: node.subtree_end,
        })
    }
    pub fn rule_name(&self, rule: CaptureRule) -> &str {
        match rule {
            CaptureRule::User(id) => &self.grammar.syntax().rules[id].name,
            CaptureRule::EndOfInput => "EOI",
        }
    }
    pub fn text(&self, id: NodeId) -> Option<&str> {
        if id.document != self.document || id.revision != self.revision {
            return None;
        }
        let span = self.nodes.get(id.index)?.span;
        Some(&self.source()[span.start..span.end])
    }
    pub fn error_position(&self, encoding: Encoding) -> Option<Position> {
        let byte = self.error.as_ref()?.offset;
        let index = self.index();
        match index.position(byte, encoding) {
            Ok(position) => Some(position),
            Err(PositionError::InsideLineBreak) => index.position(byte - 1, encoding).ok(),
            _ => None,
        }
    }
}

pub struct Session {
    document: u64,
    grammar: Arc<CompiledGrammar>,
    entry: String,
    source: OwnedIndex,
    revision: u64,
    options: MatchOptions,
}
impl Session {
    pub fn new(
        grammar: Arc<CompiledGrammar>,
        entry: impl Into<String>,
        source: Arc<str>,
        options: MatchOptions,
    ) -> Self {
        Self {
            document: NEXT_DOCUMENT
                .fetch_update(Ordering::Relaxed, Ordering::Relaxed, |id| id.checked_add(1))
                .expect("document identity space exhausted"),
            grammar,
            entry: entry.into(),
            source: OwnedIndex::new(source),
            revision: 0,
            options,
        }
    }
    pub fn revision(&self) -> u64 {
        self.revision
    }
    pub fn source(&self) -> &str {
        self.source.source()
    }
    /// Validate all coordinates/version before replacing the immutable source.
    /// This initially rebuilds the index. Incremental index/tree reuse is a
    /// later backend; this path remains its full-reparse reference baseline.
    pub fn edit(
        &mut self,
        expected_revision: u64,
        range: TextRange,
        encoding: Encoding,
        replacement: &str,
    ) -> Result<u64, EditError> {
        if expected_revision != self.revision {
            return Err(EditError::StaleRevision);
        }
        let index = self.source.index();
        let start = index
            .offset(range.start, encoding)
            .map_err(EditError::Position)?;
        let end = index
            .offset(range.end, encoding)
            .map_err(EditError::Position)?;
        if end < start {
            return Err(EditError::ReversedRange);
        }
        let revision = self
            .revision
            .checked_add(1)
            .ok_or(EditError::RevisionOverflow)?;
        let mut source =
            String::with_capacity(self.source().len() - (end - start) + replacement.len());
        source.push_str(&self.source()[..start]);
        source.push_str(replacement);
        source.push_str(&self.source()[end..]);
        self.source = OwnedIndex::new(source.into());
        self.revision = revision;
        Ok(revision)
    }
    pub fn parse(&self) -> Report {
        let (nodes, consumed, error) =
            match self
                .grammar
                .parse_source_with_options(&self.entry, self.source(), self.options)
            {
                Ok(tree) if tree.consumed == self.source().len() => {
                    (Arc::from(tree.nodes), Some(tree.consumed), None)
                }
                Ok(tree) => (
                    Arc::from([]),
                    None,
                    Some(ParseError {
                        kind: crate::ParseErrorKind::Syntax,
                        offset: tree.consumed,
                        expected: vec!["EOI".into()],
                        message: "trailing source after editor entry".into(),
                    }),
                ),
                Err(error) => (Arc::from([]), None, Some(error)),
            };
        Report {
            document: self.document,
            revision: self.revision,
            grammar: self.grammar.clone(),
            source: self.source.clone(),
            nodes,
            consumed,
            error,
        }
    }
}
