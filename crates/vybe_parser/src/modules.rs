//! In-memory grammar modules. Definitions retain ordinary grammar syntax;
//! imports/exports are explicit Rust metadata, ready for a declarative loader.
//! Scoped trivia belongs to the called rule's module, never to a flattened
//! global WHITESPACE definition.
use crate::{Builtin, CompiledGrammar, Diagnostic, ExprKind, GrammarSyntax, Span, program::Scope};
use std::collections::{HashMap, HashSet};

pub struct Import<'a> {
    pub local: &'a str,
    pub module: &'a str,
    pub rule: &'a str,
}
pub struct Module<'a> {
    pub name: &'a str,
    pub source: &'a str,
    pub exports: &'a [&'a str],
    pub imports: &'a [Import<'a>],
}
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Location {
    pub module: usize,
    pub span: Span,
}
#[derive(Debug, Clone)]
pub struct ModuleDiagnostic {
    pub code: &'static str,
    pub message: String,
    pub location: Location,
    pub related: Vec<(Location, String)>,
}
#[derive(Debug)]
pub struct LinkedGrammar {
    grammar: CompiledGrammar,
    source: String,
    ranges: Vec<Span>,
    names: Vec<String>,
}
impl LinkedGrammar {
    pub fn grammar(&self) -> &CompiledGrammar {
        &self.grammar
    }
    pub fn source(&self) -> &str {
        &self.source
    }
    pub fn module_name(&self, module: usize) -> &str {
        &self.names[module]
    }
    pub fn location(&self, span: Span) -> Location {
        locate(&self.ranges, span)
    }
    pub fn warnings(&self) -> Vec<ModuleDiagnostic> {
        self.grammar
            .analysis()
            .warnings
            .iter()
            .cloned()
            .map(|error| remap(&self.ranges, error))
            .collect()
    }
}
fn locate(ranges: &[Span], span: Span) -> Location {
    let module = ranges
        .partition_point(|range| range.start <= span.start)
        .saturating_sub(1);
    let range = ranges[module];
    Location {
        module,
        span: Span {
            start: span.start.saturating_sub(range.start),
            end: span.end.saturating_sub(range.start),
        },
    }
}
fn remap(ranges: &[Span], error: Diagnostic) -> ModuleDiagnostic {
    ModuleDiagnostic {
        code: error.code,
        message: error.message,
        location: locate(ranges, error.span),
        related: error
            .related
            .into_iter()
            .map(|(span, message)| (locate(ranges, span), message))
            .collect(),
    }
}
fn error(
    module: usize,
    code: &'static str,
    message: impl Into<String>,
    span: Span,
) -> ModuleDiagnostic {
    ModuleDiagnostic {
        code,
        message: message.into(),
        location: Location { module, span },
        related: Vec::new(),
    }
}
fn internal(module: usize, name: &str) -> String {
    format!("m{module}_{name}")
}
fn shifted(span: Span, offset: usize) -> Span {
    Span {
        start: span.start + offset,
        end: span.end + offset,
    }
}

pub fn compile(modules: &[Module<'_>]) -> Result<LinkedGrammar, Vec<ModuleDiagnostic>> {
    if modules.is_empty() {
        return Err(vec![error(
            0,
            "M001",
            "no grammar modules supplied",
            Span { start: 0, end: 0 },
        )]);
    }
    let mut diagnostics = Vec::new();
    let mut names = HashMap::new();
    let mut syntax = Vec::new();
    let mut locals = Vec::new();
    for (id, module) in modules.iter().enumerate() {
        if module.name.is_empty() || names.insert(module.name, id).is_some() {
            diagnostics.push(error(
                id,
                "M001",
                "empty or duplicate module name",
                Span { start: 0, end: 0 },
            ));
        }
        match crate::parse(module.source) {
            Ok(parsed) => {
                let mut local = HashMap::new();
                for (rule_id, rule) in parsed.rules.iter().enumerate() {
                    if let Some(previous) = local.insert(rule.name.clone(), rule_id) {
                        let mut duplicate = error(
                            id,
                            "G009",
                            format!("duplicate rule '{}'", rule.name),
                            rule.name_span,
                        );
                        duplicate.related.push((
                            Location {
                                module: id,
                                span: parsed.rules[previous].name_span,
                            },
                            "first definition".into(),
                        ));
                        diagnostics.push(duplicate);
                    }
                }
                syntax.push(parsed);
                locals.push(local);
            }
            Err(diagnostic) => {
                diagnostics.push(ModuleDiagnostic {
                    code: diagnostic.code,
                    message: diagnostic.message,
                    location: Location {
                        module: id,
                        span: diagnostic.span,
                    },
                    related: Vec::new(),
                });
                syntax.push(GrammarSyntax {
                    rules: Vec::new(),
                    expressions: Vec::new(),
                });
                locals.push(HashMap::new());
            }
        }
    }
    if !diagnostics.is_empty() {
        return Err(diagnostics);
    }
    let mut exports = Vec::new();
    for (id, module) in modules.iter().enumerate() {
        let mut exported = HashSet::new();
        for name in module.exports {
            if !locals[id].contains_key(*name) || !exported.insert(*name) {
                diagnostics.push(error(
                    id,
                    "M002",
                    format!("missing or duplicate export '{name}'"),
                    Span { start: 0, end: 0 },
                ));
            }
        }
        exports.push(exported);
    }
    let mut bindings: Vec<HashMap<String, String>> = locals
        .iter()
        .enumerate()
        .map(|(id, local)| {
            local
                .keys()
                .map(|name| (name.clone(), internal(id, name)))
                .collect()
        })
        .collect();
    for (id, module) in modules.iter().enumerate() {
        for import in module.imports {
            let Some(target) = names
                .get(import.module)
                .copied()
                .filter(|target| exports[*target].contains(import.rule))
            else {
                diagnostics.push(error(
                    id,
                    "M003",
                    format!(
                        "import '{}::{}' is missing or private",
                        import.module, import.rule
                    ),
                    Span { start: 0, end: 0 },
                ));
                continue;
            };
            if bindings[id]
                .insert(import.local.into(), internal(target, import.rule))
                .is_some()
            {
                diagnostics.push(error(
                    id,
                    "M004",
                    format!(
                        "import alias '{}' collides with another binding",
                        import.local
                    ),
                    Span { start: 0, end: 0 },
                ));
            }
        }
        for expression in &syntax[id].expressions {
            if let ExprKind::Reference(name) = &expression.kind
                && !bindings[id].contains_key(name)
                && Builtin::from_name(name).is_none()
            {
                diagnostics.push(error(
                    id,
                    "G010",
                    format!("undefined rule or unsupported builtin '{name}'"),
                    expression.span,
                ));
            }
        }
    }
    if !diagnostics.is_empty() {
        return Err(diagnostics);
    }
    let mut merged = GrammarSyntax {
        rules: Vec::new(),
        expressions: Vec::new(),
    };
    let mut source = String::new();
    let mut ranges = Vec::new();
    let mut rule_scopes = Vec::new();
    for (id, parsed) in syntax.into_iter().enumerate() {
        let offset = source.len();
        source.push_str(modules[id].source);
        ranges.push(Span {
            start: offset,
            end: source.len(),
        });
        source.push('\n');
        let base = merged.expressions.len();
        for mut expression in parsed.expressions {
            expression.span = shifted(expression.span, offset);
            match &mut expression.kind {
                ExprKind::Reference(name) => {
                    if let Some(bound) = bindings[id].get(name) {
                        *name = bound.clone();
                    }
                }
                ExprKind::Sequence { left, right } | ExprKind::Choice { left, right } => {
                    *left += base;
                    *right += base;
                }
                ExprKind::Repeat { expression, .. }
                | ExprKind::Predicate { expression, .. }
                | ExprKind::Group(expression)
                | ExprKind::Push(expression) => *expression += base,
                _ => {}
            }
            merged.expressions.push(expression);
        }
        for mut rule in parsed.rules {
            rule.name = internal(id, &rule.name);
            rule.name_span = shifted(rule.name_span, offset);
            rule.span = shifted(rule.span, offset);
            rule.expression += base;
            merged.rules.push(rule);
            rule_scopes.push(id);
        }
    }
    let mut grammar = crate::resolve(merged).map_err(|errors| {
        errors
            .into_iter()
            .map(|error| remap(&ranges, error))
            .collect::<Vec<_>>()
    })?;
    grammar.rule_scopes = rule_scopes;
    grammar.scopes = bindings
        .iter()
        .map(|bindings| {
            let whitespace = bindings
                .get("WHITESPACE")
                .and_then(|name| grammar.rule_id(name));
            let comment = bindings
                .get("COMMENT")
                .and_then(|name| grammar.rule_id(name));
            let (whitespace_class, whitespace_prefix_class) =
                crate::lexical::whitespace_classes(&grammar, whitespace);
            Scope {
                whitespace,
                comment,
                whitespace_class,
                whitespace_prefix_class,
            }
        })
        .collect();
    for (id, module) in modules.iter().enumerate() {
        for exported in module.exports {
            let rule = grammar.rule_id(&internal(id, exported)).unwrap();
            grammar.add_entry(format!("{}::{exported}", module.name), rule);
        }
    }
    grammar.refresh_analysis();
    if !grammar.analysis().errors.is_empty() {
        return Err(grammar
            .analysis()
            .errors
            .iter()
            .cloned()
            .map(|error| remap(&ranges, error))
            .collect());
    }
    Ok(LinkedGrammar {
        grammar,
        source,
        ranges,
        names: modules.iter().map(|module| module.name.into()).collect(),
    })
}
