//! Explicit, parser-owned specialization of an existing emitter.
//! Facts describe physically emitted operands after argument adaptation.
use vybe_runtime::registry::emitter_results::{EmitOutcome, EmitRequest};
use vybe_runtime::registry::{EmitDispatchFn, ParseFn};
use vybe_runtime::{Chunk, Op, Value};

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum ProducedValueProof {
    Unknown,
    /// Immutable string-global constant in the current caller chunk.
    StringLiteral {
        constant: u32,
    },
}

/// An opt-in sibling of the selected legacy emitter, never a name resolver.
/// Unhandled must leave all chunks unchanged. The descriptor may handle only
/// operations dispatched by its legacy producer, preserving its stack ABI.
/// Only Value/DiscardedValue requests are supported; typed result-position
/// providers retain their separate dispatch and representation authority.
pub struct Definition {
    pub operation: &'static str,
    pub parser: ParseFn,
    pub producer: EmitDispatchFn,
    pub dispatch: fn(
        &str,
        &mut Vec<Chunk>,
        usize,
        u8,
        u32,
        EmitRequest,
        &[ProducedValueProof],
    ) -> EmitOutcome,
}

#[derive(Clone, Copy, Default)]
pub struct Selection {
    parser: Option<ParseFn>,
    definitions: &'static [Definition],
}
impl Selection {
    pub fn new(parser: ParseFn, definitions: &'static [Definition]) -> Self {
        Self {
            parser: Some(parser),
            definitions,
        }
    }
    pub(crate) fn matches_parser(self, parser: ParseFn) -> bool {
        self.parser
            .is_some_and(|selected| std::ptr::fn_addr_eq(selected, parser))
    }
    pub(crate) fn resolve(self, name: &str) -> Option<&'static Definition> {
        if self.definitions.is_empty() {
            return None;
        }
        let parser = self.parser?;
        // Do not resolve a namespace for unrelated common operations. Keep the
        // first valid descriptor order when a namespace has several producers.
        let mut candidates = self
            .definitions
            .iter()
            .filter(|d| d.operation == name && std::ptr::fn_addr_eq(d.parser, parser));
        let first = candidates.next()?;
        let (namespace, _) = name.split_once('.')?;
        let producer = crate::languages::emit_dispatch_for(namespace)?;
        std::iter::once(first)
            .chain(candidates)
            .find(|d| std::ptr::fn_addr_eq(d.producer, producer))
    }
}

/// Inspect the actual just-emitted operand; do not intern/reorder constants.
/// Extra coercion, address wrappers, nested calls and inferred types decline.
pub(crate) fn string_literal_proof(chunk: &Chunk, start: usize, text: &str) -> ProducedValueProof {
    let Some(bytes) = chunk.code.get(start..) else {
        return ProducedValueProof::Unknown;
    };
    if bytes.len() != 8 || Op::from_code_at(bytes, 0) != Op::GLOBAL_GET {
        return ProducedValueProof::Unknown;
    }
    let constant = u32::from_be_bytes(bytes[4..8].try_into().unwrap());
    let key = vybe_runtime::chunk::imported_global_key(
        vybe_runtime::chunk::STRING_CONSTANTS_MODULE,
        text,
    );
    if !matches!(chunk.constants.get(constant as usize), Some(Value::String(s)) if s.as_ref() == key)
        || !chunk
            .global_imports
            .iter()
            .any(|i| i.module == vybe_runtime::chunk::STRING_CONSTANTS_MODULE && i.name == text)
    {
        return ProducedValueProof::Unknown;
    }
    ProducedValueProof::StringLiteral { constant }
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn facts_require_exact_physical_producer_and_import_authority() {
        let mut c = Chunk::new("proof");
        c.emit_string_const("literal", 7);
        assert!(matches!(
            string_literal_proof(&c, 0, "literal"),
            ProducedValueProof::StringLiteral { .. }
        ));
        assert_eq!(
            string_literal_proof(&c, 0, "different"),
            ProducedValueProof::Unknown
        );
        let imports = std::mem::take(&mut c.global_imports);
        assert_eq!(
            string_literal_proof(&c, 0, "literal"),
            ProducedValueProof::Unknown
        );
        c.global_imports = imports;
        c.emit_op(Op::DROP, 7);
        assert_eq!(
            string_literal_proof(&c, 0, "literal"),
            ProducedValueProof::Unknown
        );
        assert_eq!(
            string_literal_proof(&c, usize::MAX, "literal"),
            ProducedValueProof::Unknown
        );
    }

    #[test]
    fn prior_caller_code_is_not_part_of_the_operand_and_proof_does_not_mutate() {
        let mut c = Chunk::new("caller");
        c.emit_string_const("prior", 2);
        let start = c.code.len();
        c.emit_string_const("\0Unicode: café😀", 9);
        let before = format!("{c:?}");
        assert!(matches!(
            string_literal_proof(&c, start, "\0Unicode: café😀"),
            ProducedValueProof::StringLiteral { .. }
        ));
        assert_eq!(format!("{c:?}"), before);
        assert_eq!(
            string_literal_proof(&c, 0, "prior"),
            ProducedValueProof::Unknown
        );
    }

    #[test]
    fn default_selection_never_resolves_a_registered_namespace() {
        assert!(Selection::default().resolve("php.str_replace").is_none());
    }
}
