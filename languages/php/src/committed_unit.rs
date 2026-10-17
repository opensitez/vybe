//! Opt-in typed source-unit handoff to a reusable compiler session.
//! PHP currently constructs expression recipes in Grameo and completes scoped
//! statement/class normalization afterward. This API preserves that work; it
//! does not advertise statement AST or VM publication before source-unit EOF.
use grameo::source_units::Request;
use vybe_ast::Module;

/// Receipt from the actual successful Grameo Builder::committed callback.
/// Byte counts address normalized parser input, which can differ from mixed
/// HTML source. Construction is private to the validated frontend path.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct GrammarCommit {
    source_bytes: usize,
    parser_bytes: usize,
    consumed: usize,
}
impl GrammarCommit {
    pub(crate) fn validated(
        source_bytes: usize,
        parser_bytes: usize,
        consumed: usize,
    ) -> Result<Self, String> {
        if consumed != parser_bytes {
            return Err("PHP source-unit entry did not validate full parser input".into());
        }
        Ok(Self {
            source_bytes,
            parser_bytes,
            consumed,
        })
    }
    pub fn source_bytes(self) -> usize {
        self.source_bytes
    }
    pub fn parser_bytes(self) -> usize {
        self.parser_bytes
    }
    pub fn consumed(self) -> usize {
        self.consumed
    }
}

/// Complete normalized AST/manifest of one include file or REPL submission.
/// Contains existing scope/class/slot/directive, imports, late declarations,
/// source intents and contextual normalization without a second parse.
/// It does not install declarations, execute includes or publish VM chunks.
pub struct CommittedSourceUnit {
    module: Module,
    source_units: Vec<Request<String>>,
    grammar_commit: GrammarCommit,
}
impl CommittedSourceUnit {
    pub fn manifest(&self) -> &Module {
        &self.module
    }
    pub fn source_units(&self) -> &[Request<String>] {
        &self.source_units
    }
    pub fn grammar_commit(&self) -> GrammarCommit {
        self.grammar_commit
    }
    /// Move the original AST into Compiler::begin_emission_owned; no statement
    /// clone, per-statement Module compilation or whole-request retention.
    pub fn into_parts(self) -> (Module, Vec<Request<String>>, GrammarCommit) {
        (self.module, self.source_units, self.grammar_commit)
    }
}

/// Parse and normalize one source unit once, then invoke the typed consumer
/// exactly once. Grammar/semantic failure never invokes it. The consumer may
/// begin one PreparedEmission session, emit complete functions/statements and
/// finalize once. Its errors are returned unchanged; unpublished state belongs
/// to that session and must be dropped/aborted, not installed in the VM.
/// Ordinary parse/PEG entry points are unchanged and never call this consumer.
pub fn parse_committed_source_unit<T>(
    source: &str,
    consume: impl FnOnce(CommittedSourceUnit) -> Result<T, String>,
) -> Result<T, String> {
    let (module, source_units, grammar_commit) =
        crate::walker::parse_direct_with_committed_source_unit(source)?;
    consume(CommittedSourceUnit {
        module,
        source_units,
        grammar_commit,
    })
}
