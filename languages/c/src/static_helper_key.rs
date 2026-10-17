//! Span-independent keys for the proven, non-escaping C helper subset.
//! This is not function deduplication: the caller first proves that the
//! expression is a capture-free immediate call with a plain parameter ABI.

use std::collections::hash_map::DefaultHasher;
use std::hash::{Hash, Hasher};
use vybe_ast::{ExprKind, Expression, Literal, ObjectProperty, Param, PassBy};

pub(super) fn fingerprint(params: &[Param], body: &Expression) -> Option<u64> {
    let mut hash = DefaultHasher::new();
    params.len().hash(&mut hash);
    for (index, param) in params.iter().enumerate() {
        if param.type_hint.is_some() || param.default.is_some() || param.is_rest
            || param.is_kwargs || param.is_optional || param.pass_by != PassBy::Value
            || params[..index].iter().any(|prior| prior.name == param.name)
        {
            return None;
        }
        param.is_nullable.hash(&mut hash);
    }
    hash_expr(body, params, &mut hash)?;
    Some(hash.finish())
}

fn hash_expr(value: &Expression, params: &[Param], hash: &mut impl Hasher) -> Option<()> {
    std::mem::discriminant(&value.kind).hash(hash);
    match &value.kind {
        ExprKind::Lit(literal) => {
            std::mem::discriminant(literal).hash(hash);
            match literal {
                Literal::Int(value) | Literal::BigInt(value) => value.hash(hash),
                Literal::Float(value) => value.to_bits().hash(hash),
                Literal::Str(value) => value.hash(hash),
                Literal::Bool(value) => value.hash(hash),
                Literal::Char(value) => value.hash(hash),
                Literal::Bytes(value) => value.hash(hash),
                Literal::Null | Literal::Undefined | Literal::Ellipsis => {}
            }
        }
        ExprKind::Ident(name) => {
            let position = params.iter().position(|param| param.name == *name);
            position.is_some().hash(hash);
            if let Some(position) = position { position.hash(hash); }
            else { name.hash(hash); }
        }
        ExprKind::Member { object, field, null_safe } => {
            field.hash(hash);
            null_safe.hash(hash);
            hash_expr(object, params, hash)?;
        }
        ExprKind::Index { object, index, null_safe } => {
            null_safe.hash(hash);
            hash_expr(object, params, hash)?;
            hash_expr(index, params, hash)?;
        }
        ExprKind::Unary { op, expr } => {
            std::mem::discriminant(op).hash(hash);
            hash_expr(expr, params, hash)?;
        }
        ExprKind::Binary { op, left, right } => {
            // Payloads (e.g. integer width) may collide here; equivalent()
            // compares the complete operator before allowing any reuse.
            std::mem::discriminant(op).hash(hash);
            hash_expr(left, params, hash)?;
            hash_expr(right, params, hash)?;
        }
        ExprKind::Cast { expr, type_name } | ExprKind::IsType { expr, type_name } => {
            type_name.hash(hash);
            hash_expr(expr, params, hash)?;
        }
        ExprKind::Ternary { cond, then, else_ } => {
            hash_expr(cond, params, hash)?;
            hash_expr(then, params, hash)?;
            hash_expr(else_, params, hash)?;
        }
        ExprKind::Sequence(values) => {
            values.len().hash(hash);
            for value in values { hash_expr(value, params, hash)?; }
        }
        ExprKind::Object(properties) => {
            properties.len().hash(hash);
            for property in properties {
                std::mem::discriminant(property).hash(hash);
                match property {
                    ObjectProperty::KeyValue { key, value }
                    | ObjectProperty::Computed { key, value } => {
                        hash_expr(key, params, hash)?;
                        hash_expr(value, params, hash)?;
                    }
                    _ => return None,
                }
            }
        }
        ExprKind::Call { callee, args, optional: false } => {
            hash_expr(callee, params, hash)?;
            args.len().hash(hash);
            for arg in args {
                if arg.spread || arg.by_ref || arg.name.is_some() { return None; }
                hash_expr(&arg.value, params, hash)?;
            }
        }
        _ => return None,
    }
    Some(())
}

pub(super) fn equivalent(a: &[Param], left: &Expression, b: &[Param], right: &Expression) -> bool {
    a.len() == b.len() && a.iter().zip(b).all(|(a, b)| {
        a.is_nullable == b.is_nullable
    }) && same_expr(left, right, a, b)
}

fn same_literal(a: &Literal, b: &Literal) -> bool {
    match (a, b) {
        (Literal::Int(a), Literal::Int(b)) | (Literal::BigInt(a), Literal::BigInt(b)) => a == b,
        (Literal::Float(a), Literal::Float(b)) => a.to_bits() == b.to_bits(),
        (Literal::Str(a), Literal::Str(b)) => a == b,
        (Literal::Bool(a), Literal::Bool(b)) => a == b,
        (Literal::Char(a), Literal::Char(b)) => a == b,
        (Literal::Bytes(a), Literal::Bytes(b)) => a == b,
        (Literal::Null, Literal::Null) | (Literal::Undefined, Literal::Undefined)
        | (Literal::Ellipsis, Literal::Ellipsis) => true,
        _ => false,
    }
}

fn same_expr(a: &Expression, b: &Expression, params: &[Param], other: &[Param]) -> bool {
    let same = |a: &Expression, b: &Expression| same_expr(a, b, params, other);
    match (&a.kind, &b.kind) {
        (ExprKind::Lit(a), ExprKind::Lit(b)) => same_literal(a, b),
        (ExprKind::Ident(a), ExprKind::Ident(b)) => {
            match (params.iter().position(|param| param.name == *a),
                   other.iter().position(|param| param.name == *b)) {
                (Some(a), Some(b)) => a == b,
                (None, None) => a == b,
                _ => false,
            }
        }
        (ExprKind::Member { object: a, field: af, null_safe: an },
         ExprKind::Member { object: b, field: bf, null_safe: bn }) => af == bf && an == bn && same(a, b),
        (ExprKind::Index { object: a, index: ai, null_safe: an },
         ExprKind::Index { object: b, index: bi, null_safe: bn }) => an == bn && same(a, b) && same(ai, bi),
        (ExprKind::Unary { op: ao, expr: a }, ExprKind::Unary { op: bo, expr: b }) => ao == bo && same(a, b),
        (ExprKind::Binary { op: ao, left: a, right: ar },
         ExprKind::Binary { op: bo, left: b, right: br }) => ao == bo && same(a, b) && same(ar, br),
        (ExprKind::Cast { expr: a, type_name: at }, ExprKind::Cast { expr: b, type_name: bt })
        | (ExprKind::IsType { expr: a, type_name: at }, ExprKind::IsType { expr: b, type_name: bt }) => at == bt && same(a, b),
        (ExprKind::Ternary { cond: a, then: at, else_: ae },
         ExprKind::Ternary { cond: b, then: bt, else_: be }) => same(a, b) && same(at, bt) && same(ae, be),
        (ExprKind::Sequence(a), ExprKind::Sequence(b)) => a.len() == b.len() && a.iter().zip(b).all(|(a, b)| same(a, b)),
        (ExprKind::Object(a), ExprKind::Object(b)) => a.len() == b.len() && a.iter().zip(b).all(|(a, b)| match (a, b) {
            (ObjectProperty::KeyValue { key: a, value: av }, ObjectProperty::KeyValue { key: b, value: bv })
            | (ObjectProperty::Computed { key: a, value: av }, ObjectProperty::Computed { key: b, value: bv }) => same(a, b) && same(av, bv),
            _ => false,
        }),
        (ExprKind::Call { callee: a, args: aa, optional: false },
         ExprKind::Call { callee: b, args: ba, optional: false }) => same(a, b) && aa.len() == ba.len()
            && aa.iter().zip(ba).all(|(a, b)| !a.spread && !b.spread && !a.by_ref && !b.by_ref
                && a.name.is_none() && b.name.is_none() && same(&a.value, &b.value)),
        _ => false,
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use vybe_ast::{BinOp, BitLane, IntOp};

    fn param(name: &str) -> Param {
        Param { name: name.into(), type_hint: None, default: None,
            pass_by: PassBy::Value, is_rest: false, is_kwargs: false,
            is_optional: false, is_nullable: false }
    }

    #[test]
    fn parameter_names_are_alpha_equivalent_but_positions_and_free_names_are_not() {
        let a = [param("left"), param("right")];
        let b = [param("first"), param("second")];
        let make = |left: &str, right: &str| Expression::new(ExprKind::Binary {
            op: BinOp::Integer(IntOp::Sub, BitLane::W32),
            left: Box::new(Expression::ident(left)), right: Box::new(Expression::ident(right)),
        });
        let left = make("left", "right");
        let right = make("first", "second");
        assert_eq!(fingerprint(&a, &left), fingerprint(&b, &right));
        assert!(equivalent(&a, &left, &b, &right));
        assert!(!equivalent(&a, &left, &b, &make("second", "first")));
        assert!(!equivalent(&a, &Expression::ident("left"), &b, &Expression::ident("left")));
        assert!(equivalent(&a, &Expression::ident("support"), &b, &Expression::ident("support")));
        let mut nullable = b.clone();
        nullable[0].is_nullable = true;
        assert!(!equivalent(&a, &left, &nullable, &right));
        assert!(fingerprint(&[param("x"), param("x")], &Expression::ident("x")).is_none());
    }

    #[test]
    fn collision_checks_keep_integer_lanes_distinct() {
        let make = |lane| Expression::new(ExprKind::Binary {
            op: BinOp::Integer(IntOp::Add, lane),
            left: Box::new(Expression::int(1)), right: Box::new(Expression::int(2)),
        });
        let a = make(BitLane::W32);
        let b = make(BitLane::W64);
        assert_eq!(fingerprint(&[], &a), fingerprint(&[], &b));
        assert!(!equivalent(&[], &a, &[], &b));
        assert!(equivalent(&[], &a, &[], &a));
    }

    #[test]
    fn preserves_float_bits_and_rejects_unknown_nodes() {
        let make = |value| Expression::new(ExprKind::Lit(Literal::Float(value)));
        assert!(!equivalent(&[], &make(0.0), &[], &make(-0.0)));
        let nan = make(f64::from_bits(0x7ff8000000000001));
        assert!(equivalent(&[], &nan, &[], &nan));
        assert!(!equivalent(&[], &nan, &[], &make(f64::from_bits(0x7ff8000000000002))));
        assert!(fingerprint(&[], &Expression::new(ExprKind::This)).is_none());
    }
}
