//! PHP numeric literal semantics shared by capture and direct AST bindings.
use vybe_ast::{ExprKind, Expression, Literal};

pub(super) fn number(raw: &str) -> Expression {
    let s = if raw.as_bytes().contains(&b'_') {
        std::borrow::Cow::Owned(raw.replace('_', ""))
    } else {
        std::borrow::Cow::Borrowed(raw)
    };
    let kind = if let Some(rest) = s.strip_prefix("0x").or_else(|| s.strip_prefix("0X")) {
        i64::from_str_radix(rest, 16)
            .map(Literal::Int)
            .map(ExprKind::Lit)
            .unwrap_or(ExprKind::Lit(Literal::Int(0)))
    } else if let Some(rest) = s.strip_prefix("0b").or_else(|| s.strip_prefix("0B")) {
        i64::from_str_radix(rest, 2)
            .map(Literal::Int)
            .map(ExprKind::Lit)
            .unwrap_or(ExprKind::Lit(Literal::Int(0)))
    } else if let Some(rest) = s.strip_prefix("0o").or_else(|| s.strip_prefix("0O")) {
        i64::from_str_radix(rest, 8)
            .map(Literal::Int)
            .map(ExprKind::Lit)
            .unwrap_or(ExprKind::Lit(Literal::Int(0)))
    } else if s.len() > 1 && s.starts_with('0') && s.bytes().all(|c| (b'0'..=b'7').contains(&c)) {
        i64::from_str_radix(&s[1..], 8)
            .map(Literal::Int)
            .map(ExprKind::Lit)
            .unwrap_or(ExprKind::Lit(Literal::Int(0)))
    } else if s.contains('.') || s.contains('e') || s.contains('E') {
        s.parse::<f64>()
            .map(Literal::Float)
            .map(ExprKind::Lit)
            .unwrap_or(ExprKind::Lit(Literal::Int(0)))
    } else {
        match s.parse::<i128>() {
            Ok(n) if n == (i64::MAX as i128) + 1 => ExprKind::Lit(Literal::BigInt(i64::MIN)),
            Ok(n) if n > i64::MAX as i128 || n < i64::MIN as i128 => {
                ExprKind::Lit(Literal::Float(n as f64))
            }
            Ok(n) if n.abs() > 9_007_199_254_740_991_i128 => {
                ExprKind::Lit(Literal::BigInt(n as i64))
            }
            Ok(n) => ExprKind::Lit(Literal::Int(n as i64)),
            Err(_) => ExprKind::Lit(Literal::Int(0)),
        }
    };
    Expression::new(kind)
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn php_numeric_boundaries_and_bases() {
        for (input, expected) in [
            ("0", 0),
            ("17", 17),
            ("1_024", 1024),
            ("0xff", 255),
            ("0XFF", 255),
            ("0b1010", 10),
            ("0B1010", 10),
            ("0o77", 63),
            ("0O77", 63),
            ("077", 63),
            ("0xFFFF_FFFF", 4294967295),
        ] {
            assert!(
                matches!(number(input).kind,ExprKind::Lit(Literal::Int(n)) if n==expected),
                "{input}"
            );
        }
        for (input, expected) in [
            ("9007199254740992", 9007199254740992),
            ("9223372036854775807", i64::MAX),
            ("9223372036854775808", i64::MIN),
        ] {
            assert!(
                matches!(number(input).kind,ExprKind::Lit(Literal::BigInt(n)) if n==expected),
                "{input}"
            );
        }
        for (input, expected) in [
            ("1.25", 1.25),
            ("1_2.5", 12.5),
            ("1e2", 100.0),
            ("1E-2", 0.01),
            ("9223372036854775809", 9223372036854775809_i128 as f64),
        ] {
            assert!(
                matches!(number(input).kind,ExprKind::Lit(Literal::Float(n)) if n==expected),
                "{input}"
            );
        }
    }
}
