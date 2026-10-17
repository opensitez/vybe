use grameo::{
    CompiledGrammar, Diagnostic,
    ast::{Cardinality, SourceSyntax},
};
use std::fmt::Write;

pub(super) fn emit(
    grammar: &CompiledGrammar,
    out: &mut String,
    entry: bool,
    schema: Option<&[grameo::ast::ConstructorSchema<'_>]>,
) -> Result<(), Vec<Diagnostic>> {
    let bindings = &grammar.syntax().ast_bindings;
    if bindings.is_empty() {
        return Ok(());
    }
    let mut constructors = bindings
        .iter()
        .map(|b| b.constructor.as_str())
        .collect::<Vec<_>>();
    constructors.sort_unstable();
    constructors.dedup();
    let errors = bindings
        .iter()
        .filter(|b| {
            matches!(
                b.constructor.as_str(),
                "self" | "Self" | "super" | "crate" | "_"
            )
        })
        .map(|b| {
            Diagnostic::new(
                "G043",
                "AST constructor cannot be represented as a Rust identifier",
                b.span,
            )
        })
        .collect::<Vec<_>>();
    if !errors.is_empty() {
        return Err(errors);
    }
    out.push_str("#[derive(Debug, Clone, Copy, PartialEq, Eq)]\n#[allow(non_camel_case_types)]\n#[repr(usize)]\npub enum AstConstructor {\n");
    for name in &constructors {
        let id = schema
            .map(|schema| schema.iter().position(|c| c.name == *name).unwrap())
            .unwrap_or_else(|| constructors.binary_search(name).unwrap());
        writeln!(out, "    r#{name} = {id},").unwrap();
    }
    out.push_str("}\npub static AST_BINDINGS: &[::grameo::ast::Binding] = &[\n");
    for binding in bindings {
        let rule = grammar.rule_id(&binding.rule).unwrap();
        let constructor = schema
            .map(|schema| {
                schema
                    .iter()
                    .position(|c| c.name == binding.constructor)
                    .unwrap()
            })
            .unwrap_or_else(|| {
                constructors
                    .binary_search(&binding.constructor.as_str())
                    .unwrap()
            });
        writeln!(
            out,
            "    ::grameo::ast::Binding {{ rule: {rule}, constructor: {constructor}, fields: &["
        )
        .unwrap();
        for field in &binding.fields {
            let source = match &field.source {
                SourceSyntax::Children(cardinality) => format!(
                    "::grameo::ast::Source::Children(::grameo::ast::Cardinality::{cardinality:?})"
                ),
                SourceSyntax::Text => "::grameo::ast::Source::Text".to_owned(),
                SourceSyntax::Span => "::grameo::ast::Source::Span".to_owned(),
                SourceSyntax::Child {
                    rule,
                    occurrence,
                    cardinality,
                } => {
                    let rule = grammar.rule_id(rule).unwrap();
                    let cardinality = match cardinality {
                        Cardinality::One => "One",
                        Cardinality::Optional => "Optional",
                        Cardinality::Many => "Many",
                    };
                    format!(
                        "::grameo::ast::Source::Child {{ rule: {rule}, occurrence: {occurrence:?}, cardinality: ::grameo::ast::Cardinality::{cardinality} }}"
                    )
                }
            };
            writeln!(
                out,
                "        ::grameo::ast::Field {{ name: {:?}, source: {source} }},",
                field.name
            )
            .unwrap();
        }
        out.push_str("    ] },\n");
    }
    out.push_str("] ;\n");
    if entry {
        out.push_str("impl Parser {\n    pub fn build_bound_ast<T: ::grameo::ast::builder::Target>(&self, rule: Rule, input: &str) -> Result<T::Value, ::grameo::ParseError> {\n        self.build_bound_ast_with_context::<T>(rule, input, T::prepare(input))\n    }\n    /// Build using the frontend's immutable per-input scope/directive policy.\n    pub fn build_bound_ast_with_context<T: ::grameo::ast::builder::Target>(&self, rule: Rule, input: &str, context: T::Context) -> Result<T::Value, ::grameo::ParseError> {\n        let mut builder = ::grameo::ast::builder::AstBuilder::<T>::with_context(AST_BINDINGS, context);\n        self.build_ast(rule, input, &mut builder)?;\n        builder.into_value().map_err(Into::into)\n    }\n}\n");
    }
    Ok(())
}
