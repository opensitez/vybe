use crate::lexer::{self, Kind, Token};
use crate::{Diagnostic, Span};

pub type ExprId = usize;
pub type RuleId = usize;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum RuleMode {
    Normal,
    Silent,
    Atomic,
    CompoundAtomic,
    NonAtomic,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Rule {
    pub name: String,
    pub name_span: Span,
    pub span: Span,
    pub mode: RuleMode,
    pub expression: ExprId,
}

/// Children always precede their parent in the contiguous expression arena.
/// Repetitions retain bounds; a billion repetitions do not expand the IR.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ExprKind {
    Literal {
        text: String,
        insensitive: bool,
    },
    Range {
        start: char,
        end: char,
    },
    Reference(String),
    Sequence {
        left: ExprId,
        right: ExprId,
    },
    Choice {
        left: ExprId,
        right: ExprId,
    },
    Repeat {
        expression: ExprId,
        min: u32,
        max: Option<u32>,
    },
    Predicate {
        expression: ExprId,
        positive: bool,
    },
    Group(ExprId),
    Push(ExprId),
    PushLiteral(String),
    PeekSlice {
        start: i32,
        end: Option<i32>,
    },
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Expr {
    pub kind: ExprKind,
    pub span: Span,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct GrammarSyntax {
    pub rules: Vec<Rule>,
    pub expressions: Vec<Expr>,
}

#[derive(Debug, Clone, Copy)]
pub struct ParseOptions {
    pub max_nesting: usize,
}
impl Default for ParseOptions {
    fn default() -> Self {
        Self { max_nesting: 128 }
    }
}

pub fn parse(source: &str) -> Result<GrammarSyntax, Diagnostic> {
    parse_with_options(source, ParseOptions::default())
}

pub fn parse_with_options(
    source: &str,
    options: ParseOptions,
) -> Result<GrammarSyntax, Diagnostic> {
    let mut parser = Parser {
        tokens: lexer::lex(source)?,
        position: 0,
        expressions: Vec::new(),
        depth: 0,
        options,
    };
    let mut rules = Vec::new();
    while parser.current().kind != Kind::Eof {
        rules.push(parser.rule()?);
    }
    if rules.is_empty() {
        return Err(parser.error("G005", "expected at least one rule definition"));
    }
    Ok(GrammarSyntax {
        rules,
        expressions: parser.expressions,
    })
}

struct Parser {
    tokens: Vec<Token>,
    position: usize,
    expressions: Vec<Expr>,
    depth: usize,
    options: ParseOptions,
}

impl Parser {
    fn current(&self) -> &Token {
        &self.tokens[self.position]
    }
    fn advance(&mut self) -> Token {
        let token = self.current().clone();
        if token.kind != Kind::Eof {
            self.position += 1;
        }
        token
    }
    fn error(&self, code: &'static str, message: impl Into<String>) -> Diagnostic {
        Diagnostic::new(code, message, self.current().span)
    }
    fn eat(&mut self, symbol: char) -> bool {
        if self.current().kind == Kind::Symbol(symbol) {
            self.advance();
            true
        } else {
            false
        }
    }
    fn expect(&mut self, symbol: char) -> Result<Token, Diagnostic> {
        if self.current().kind == Kind::Symbol(symbol) {
            Ok(self.advance())
        } else {
            Err(self.error("G005", format!("expected '{symbol}'")))
        }
    }
    fn add(&mut self, kind: ExprKind, span: Span) -> ExprId {
        let id = self.expressions.len();
        self.expressions.push(Expr { kind, span });
        id
    }
    fn rule(&mut self) -> Result<Rule, Diagnostic> {
        let token = self.advance();
        let Kind::Ident(name) = token.kind else {
            return Err(Diagnostic::new("G005", "expected a rule name", token.span));
        };
        self.expect('=')?;
        let mode = match self.current().kind {
            Kind::Ident(ref text) if text == "_" => {
                self.advance();
                RuleMode::Silent
            }
            Kind::Symbol('@') => {
                self.advance();
                RuleMode::Atomic
            }
            Kind::Symbol('$') => {
                self.advance();
                RuleMode::CompoundAtomic
            }
            Kind::Symbol('!') => {
                self.advance();
                RuleMode::NonAtomic
            }
            _ => RuleMode::Normal,
        };
        self.expect('{')?;
        let expression = self.expression(0)?;
        let end = self.expect('}')?.span.end;
        Ok(Rule {
            name,
            name_span: token.span,
            span: Span {
                start: token.span.start,
                end,
            },
            mode,
            expression,
        })
    }
    // Pratt parsing of the grammar language itself: postfix > predicate >
    // sequence > choice. This does not reinterpret source-language precedence.
    fn expression(&mut self, min_bp: u8) -> Result<ExprId, Diagnostic> {
        if self.depth >= self.options.max_nesting {
            return Err(self.error("G006", "grammar nesting limit exceeded"));
        }
        self.depth += 1;
        let result = self.expression_inner(min_bp);
        self.depth -= 1;
        result
    }
    fn expression_inner(&mut self, min_bp: u8) -> Result<ExprId, Diagnostic> {
        let mut left = self.primary()?;
        while let Kind::Symbol(symbol) = self.current().kind {
            let bp = match symbol {
                '*' | '+' | '?' | '{' => 40,
                '~' => 20,
                '|' => 10,
                _ => break,
            };
            if bp < min_bp {
                break;
            }
            self.advance();
            let start = self.expressions[left].span.start;
            let kind = match symbol {
                '*' => ExprKind::Repeat {
                    expression: left,
                    min: 0,
                    max: None,
                },
                '+' => ExprKind::Repeat {
                    expression: left,
                    min: 1,
                    max: None,
                },
                '?' => ExprKind::Repeat {
                    expression: left,
                    min: 0,
                    max: Some(1),
                },
                '{' => {
                    let (min, max) = self.bounds()?;
                    ExprKind::Repeat {
                        expression: left,
                        min,
                        max,
                    }
                }
                '~' | '|' => {
                    let right = self.expression(bp + 1)?;
                    if symbol == '~' {
                        ExprKind::Sequence { left, right }
                    } else {
                        ExprKind::Choice { left, right }
                    }
                }
                _ => unreachable!(),
            };
            let end = self.tokens[self.position - 1].span.end;
            left = self.add(kind, Span { start, end });
        }
        Ok(left)
    }
    fn primary(&mut self) -> Result<ExprId, Diagnostic> {
        let token = self.advance();
        let start = token.span.start;
        let kind = match token.kind {
            Kind::String(text) => ExprKind::Literal {
                text,
                insensitive: false,
            },
            Kind::Symbol('^') => {
                let literal = self.advance();
                let Kind::String(text) = literal.kind else {
                    return Err(Diagnostic::new(
                        "G005",
                        "expected a string after '^'",
                        literal.span,
                    ));
                };
                ExprKind::Literal {
                    text,
                    insensitive: true,
                }
            }
            Kind::Character(start_char) => {
                if self.current().kind != Kind::Range {
                    return Err(self.error("G005", "expected '..' after range start"));
                }
                self.advance();
                let endpoint = self.advance();
                let Kind::Character(end_char) = endpoint.kind else {
                    return Err(Diagnostic::new(
                        "G005",
                        "expected a character range endpoint",
                        endpoint.span,
                    ));
                };
                if start_char > end_char {
                    return Err(Diagnostic::new(
                        "G007",
                        "character range start exceeds end",
                        Span {
                            start,
                            end: endpoint.span.end,
                        },
                    ));
                }
                ExprKind::Range {
                    start: start_char,
                    end: end_char,
                }
            }
            Kind::Symbol('(') => {
                let child = self.expression(0)?;
                self.expect(')')?;
                ExprKind::Group(child)
            }
            Kind::Symbol('&' | '!') => {
                let expression = self.expression(30)?;
                ExprKind::Predicate {
                    expression,
                    positive: token.kind == Kind::Symbol('&'),
                }
            }
            Kind::Ident(name) if name == "PUSH" => {
                self.expect('(')?;
                let child = self.expression(0)?;
                self.expect(')')?;
                ExprKind::Push(child)
            }
            Kind::Ident(name) if name == "PUSH_LITERAL" => {
                self.expect('(')?;
                let literal = self.advance();
                let Kind::String(text) = literal.kind else {
                    return Err(Diagnostic::new(
                        "G005",
                        "PUSH_LITERAL requires a string",
                        literal.span,
                    ));
                };
                self.expect(')')?;
                ExprKind::PushLiteral(text)
            }
            Kind::Ident(name) if name == "PEEK" && self.eat('[') => {
                let begin = self.signed_index()?;
                if self.current().kind != Kind::Range {
                    return Err(self.error("G005", "expected '..' in PEEK slice"));
                }
                self.advance();
                let end = if self.current().kind == Kind::Symbol(']') {
                    None
                } else {
                    Some(self.signed_index()?)
                };
                self.expect(']')?;
                ExprKind::PeekSlice { start: begin, end }
            }
            Kind::Ident(name) => ExprKind::Reference(name),
            Kind::Symbol('#') => {
                return Err(Diagnostic::new(
                    "G008",
                    "grammar tags are not implemented yet",
                    token.span,
                ));
            }
            _ => {
                return Err(Diagnostic::new(
                    "G005",
                    "expected a grammar expression",
                    token.span,
                ));
            }
        };
        let end = self.tokens[self.position.saturating_sub(1)]
            .span
            .end
            .max(token.span.end);
        Ok(self.add(kind, Span { start, end }))
    }
    fn number(&mut self) -> Result<u32, Diagnostic> {
        let token = self.advance();
        match token.kind {
            Kind::Number(value) => Ok(value),
            _ => Err(Diagnostic::new(
                "G005",
                "expected an unsigned repetition bound",
                token.span,
            )),
        }
    }
    fn bounds(&mut self) -> Result<(u32, Option<u32>), Diagnostic> {
        let start = self.tokens[self.position - 1].span.start;
        let min = if self.current().kind == Kind::Symbol(',') {
            None
        } else {
            Some(self.number()?)
        };
        let max = if self.eat(',') {
            if self.current().kind == Kind::Symbol('}') {
                None
            } else {
                Some(self.number()?)
            }
        } else {
            min
        };
        let end = self.expect('}')?.span.end;
        if min.is_none() && max.is_none() {
            return Err(Diagnostic::new(
                "G007",
                "a repetition needs at least one bound",
                Span { start, end },
            ));
        }
        let min = min.unwrap_or(0);
        if max.is_some_and(|max| min > max) {
            return Err(Diagnostic::new(
                "G007",
                "repetition minimum exceeds maximum",
                Span { start, end },
            ));
        }
        Ok((min, max))
    }
    fn signed_index(&mut self) -> Result<i32, Diagnostic> {
        let negative = self.eat('-');
        let start = self.current().span;
        let magnitude = self.number()? as i64;
        i32::try_from(if negative { -magnitude } else { magnitude })
            .map_err(|_| Diagnostic::new("G004", "stack index exceeds i32", start))
    }
}
