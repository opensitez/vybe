//! Iterative source-expression parsing into typed values. Operator tables are
//! explicit language metadata; precedence is never guessed from PEG choices.
//! Grammar-bound tokenization/island dispatch is a separate integration layer.
use crate::Span;
use std::collections::HashMap;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Fixity {
    Prefix,
    InfixLeft,
    InfixRight,
    Postfix,
}
#[derive(Debug, Clone, Copy)]
pub struct Operator<'a> {
    pub symbol: &'a str,
    pub precedence: u16,
    pub fixity: Fixity,
}
#[derive(Debug)]
pub struct Table<'a> {
    operators: &'a [Operator<'a>],
    prefix: HashMap<&'a str, usize>,
    following: HashMap<&'a str, usize>,
}
impl<'a> Table<'a> {
    pub fn new(operators: &'a [Operator<'a>]) -> Result<Self, &'static str> {
        let mut table = Self {
            operators,
            prefix: HashMap::new(),
            following: HashMap::new(),
        };
        for (id, operator) in operators.iter().enumerate() {
            if operator.symbol.is_empty() {
                return Err("operator symbol must not be empty");
            }
            let map = if operator.fixity == Fixity::Prefix {
                &mut table.prefix
            } else {
                &mut table.following
            };
            if map.insert(operator.symbol, id).is_some() {
                return Err("ambiguous operator in the same syntactic position");
            }
        }
        Ok(table)
    }
}

#[derive(Debug, Clone, Copy)]
pub enum TokenKind<'a> {
    Atom,
    Symbol(&'a str),
    Open,
    Close,
}
#[derive(Debug, Clone, Copy)]
pub struct Token<'a> {
    pub kind: TokenKind<'a>,
    pub span: Span,
}
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Error {
    pub span: Span,
    pub message: &'static str,
}

/// Adapters construct arena handles or actual language/common AST expressions.
/// Hooks can report semantic conversion errors. Parsing restores the initial
/// journal cursor on syntax, resource and semantic errors.
pub trait Build {
    type Value;
    fn checkpoint(&self) -> usize;
    fn rollback(&mut self, mark: usize);
    fn atom(&mut self, span: Span) -> Result<Self::Value, Error>;
    fn prefix(
        &mut self,
        operator: usize,
        span: Span,
        right: Self::Value,
    ) -> Result<Self::Value, Error>;
    fn infix(
        &mut self,
        operator: usize,
        span: Span,
        left: Self::Value,
        right: Self::Value,
    ) -> Result<Self::Value, Error>;
    fn postfix(
        &mut self,
        operator: usize,
        span: Span,
        left: Self::Value,
    ) -> Result<Self::Value, Error>;
}

#[derive(Debug, Clone, Copy)]
pub struct Limits {
    pub max_tokens: usize,
    pub max_stack: usize,
}
impl Default for Limits {
    fn default() -> Self {
        Self {
            max_tokens: 1_000_000,
            max_stack: 4096,
        }
    }
}
enum Pending {
    Operator { id: usize, span: Span },
    Open { span: Span, values: usize },
}

/// Parse a complete expression token stream, with no recursive Rust calls.
/// Equal-precedence infix operators follow the incoming operator's declared
/// associativity; prefix binding can be lower than exponentiation.
pub fn parse<B: Build>(
    table: &Table<'_>,
    tokens: &[Token<'_>],
    builder: &mut B,
    limits: Limits,
) -> Result<B::Value, Error> {
    let mark = builder.checkpoint();
    let result = parse_inner(table, tokens, builder, limits);
    if result.is_err() {
        builder.rollback(mark);
    }
    result
}

fn parse_inner<B: Build>(
    table: &Table<'_>,
    tokens: &[Token<'_>],
    builder: &mut B,
    limits: Limits,
) -> Result<B::Value, Error> {
    let eof = tokens.last().map_or(0, |token| token.span.end);
    let eof = Span {
        start: eof,
        end: eof,
    };
    if tokens.len() > limits.max_tokens {
        return Err(Error {
            span: eof,
            message: "expression token limit exceeded",
        });
    }
    let mut values = Vec::new();
    let mut pending = Vec::new();
    let mut operand = true;
    for token in tokens {
        match token.kind {
            TokenKind::Atom => {
                if !operand {
                    return Err(Error {
                        span: token.span,
                        message: "expected an operator",
                    });
                }
                values.push(builder.atom(token.span)?);
                operand = false;
            }
            TokenKind::Open => {
                if !operand {
                    return Err(Error {
                        span: token.span,
                        message: "expected an operator before group",
                    });
                }
                pending.push(Pending::Open {
                    span: token.span,
                    values: values.len(),
                });
            }
            TokenKind::Close => {
                if operand {
                    return Err(Error {
                        span: token.span,
                        message: "expected an operand before closing group",
                    });
                }
                loop {
                    match pending.pop() {
                        Some(Pending::Open { values: base, .. }) => {
                            if values.len() != base + 1 {
                                return Err(Error {
                                    span: token.span,
                                    message: "invalid grouped expression",
                                });
                            }
                            break;
                        }
                        Some(op) => reduce(op, table, &mut values, builder)?,
                        None => {
                            return Err(Error {
                                span: token.span,
                                message: "unmatched closing group",
                            });
                        }
                    }
                }
            }
            TokenKind::Symbol(symbol) => {
                let id = if operand {
                    table.prefix.get(symbol)
                } else {
                    table.following.get(symbol)
                }
                .copied()
                .ok_or(Error {
                    span: token.span,
                    message: "unexpected operator in this position",
                })?;
                let incoming = table.operators[id];
                if !operand {
                    while let Some(Pending::Operator { id: top, .. }) = pending.last() {
                        let top = table.operators[*top];
                        let reduce_top = top.precedence > incoming.precedence
                            || (top.precedence == incoming.precedence
                                && incoming.fixity != Fixity::InfixRight);
                        if !reduce_top {
                            break;
                        }
                        reduce(pending.pop().unwrap(), table, &mut values, builder)?;
                    }
                }
                let op = Pending::Operator {
                    id,
                    span: token.span,
                };
                if incoming.fixity == Fixity::Postfix {
                    reduce(op, table, &mut values, builder)?;
                } else {
                    pending.push(op);
                    operand = true;
                }
            }
        }
        if pending.len() > limits.max_stack || values.len() > limits.max_stack {
            return Err(Error {
                span: token.span,
                message: "expression stack limit exceeded",
            });
        }
    }
    if operand {
        return Err(Error {
            span: eof,
            message: "expected an operand",
        });
    }
    while let Some(op) = pending.pop() {
        reduce(op, table, &mut values, builder)?;
    }
    if values.len() != 1 {
        return Err(Error {
            span: eof,
            message: "invalid expression",
        });
    }
    Ok(values.pop().unwrap())
}

fn reduce<B: Build>(
    op: Pending,
    table: &Table<'_>,
    values: &mut Vec<B::Value>,
    builder: &mut B,
) -> Result<(), Error> {
    let (id, span) = match op {
        Pending::Open { span, .. } => {
            return Err(Error {
                span,
                message: "unclosed group",
            });
        }
        Pending::Operator { id, span } => (id, span),
    };
    let right = values.pop().ok_or(Error {
        span,
        message: "missing operator operand",
    })?;
    let value = match table.operators[id].fixity {
        Fixity::Prefix => builder.prefix(id, span, right)?,
        Fixity::Postfix => builder.postfix(id, span, right)?,
        Fixity::InfixLeft | Fixity::InfixRight => {
            let left = values.pop().ok_or(Error {
                span,
                message: "missing left operand",
            })?;
            builder.infix(id, span, left, right)?
        }
    };
    values.push(value);
    Ok(())
}
