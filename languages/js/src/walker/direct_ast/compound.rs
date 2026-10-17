//! Owned logical/comparison chains with the production walker's ordering.
use super::*;
pub(super) fn is_rule(rule: usize) -> bool {
    rule == Rule::comparison as usize || rule == Rule::logical_expr as usize
}
pub(super) fn operator(text: &str) -> Option<BinOp> {
    Some(match text.trim() {
        "??" => BinOp::NullCoalesce,
        "||" => BinOp::Or,
        "&&" => BinOp::And,
        "|" => BinOp::BitOr,
        "^" => BinOp::BitXor,
        "&" => BinOp::BitAnd,
        "===" => BinOp::StrictEq,
        "!==" => BinOp::StrictNotEq,
        "==" => BinOp::Eq,
        "!=" => BinOp::NotEq,
        "<" => BinOp::Lt,
        ">" => BinOp::Gt,
        "<=" => BinOp::LtEq,
        ">=" => BinOp::GtEq,
        "instanceof" => BinOp::InstanceOf,
        "in" => BinOp::In,
        ">>>" => BinOp::UShr,
        ">>" => BinOp::Shr,
        "<<" => BinOp::Shl,
        _ => return None,
    })
}
pub(super) fn finish<const S: bool, const B: bool>(
    ast: &mut IslandAst<'_, S, B>,
    rule: usize,
    raw: grameo::Span,
    source: &str,
) {
    let values=ast.arena.values();
    if values.len()>=2 && matches!(ast.arena.get(values[values.len()-2]),Value::RegionStart(r) if *r==rule)
        && matches!(ast.arena.get(*values.last().unwrap()),Value::Ready(_,meta)
            if meta.raw.start==raw.start && meta.raw.end<=raw.end && meta.anchor.is_some()
                && source[meta.raw.end..raw.end].trim().is_empty()) {
        ast.arena.remove_value_below_top().unwrap();return;
    }
    let mut single = None;
    let mut children = Vec::new();
    while let Some(id) = ast.arena.pop() {
        if matches!(ast.arena.get(id),Value::RegionStart(r) if *r==rule) {
            break;
        }
        if let Some(previous) = single.replace(id) {
            if children.is_empty() {
                children.push(previous);
            }
            children.push(id);
        }
    }
    if children.is_empty() {
        if let Some(id) = single.filter(|id| {
            matches!(ast.arena.get(*id),Value::Ready(_,meta)
            if meta.raw.start==raw.start && meta.raw.end<=raw.end && meta.anchor.is_some()
                && source[meta.raw.end..raw.end].trim().is_empty())
        }) {
            ast.arena.push(id);
        } else {
            ast.arena.push(ast.gap.unwrap());
        }
        return;
    }
    children.reverse();
    let complete=children.first().is_some_and(|id| matches!(ast.arena.get(*id),Value::Ready(_,meta) if meta.raw.start==raw.start))
        && children.last().is_some_and(|id|matches!(ast.arena.get(*id),Value::Ready(_,meta) if meta.raw.end<=raw.end && source[meta.raw.end..raw.end].trim().is_empty()))
        && children.iter().all(|id|matches!(ast.arena.get(*id),Value::Ready(expr,meta) if meta.anchor.is_some() && !matches!(expr.kind,ExprKind::TypeOf(_))))
        && children.windows(2).all(|ids| {
            let (Value::Ready(_,left),Value::Ready(_,right))=(ast.arena.get(ids[0]),ast.arena.get(ids[1])) else {return false;};
            left.raw.end<=right.raw.start && operator(&source[left.raw.end..right.raw.start]).is_some()
        });
    if !complete {
        ast.arena.push(ast.gap.unwrap());
        return;
    }
    let mut children = children.into_iter();
    let mut left = children.next().unwrap();
    let current_rule = if rule == Rule::comparison as usize {
        Rule::comparison
    } else {
        Rule::logical_expr
    };
    let mut changed = false;
    for right in children {
        let (Value::Ready(_, lm), Value::Ready(_, rm)) =
            (ast.arena.get(left), ast.arena.get(right))
        else {
            unreachable!()
        };
        let op = operator(&source[lm.raw.end..rm.raw.start]).unwrap();
        left = ast.arena.binary_with_metadata(
            left,
            right,
            |left, right| {
                let (Value::Ready(mut left, lm), Value::Ready(mut right, rm)) = (left, right)
                else {
                    unreachable!()
                };
                let (ls, rs) = (left.span, right.span);
                // The walker's intermediate left folds keep the default span. Its
                // actual grammar child operands receive their captured rule spans.
                if !changed {
                    left.span = span(lm.anchor.unwrap());
                }
                right.span = span(rm.anchor.unwrap());
                let joined = grameo::Span {
                    start: lm.raw.start,
                    end: rm.raw.end,
                };
                (
                    Value::Ready(
                        Expression::new(ExprKind::Binary {
                            op,
                            left: Box::new(left),
                            right: Box::new(right),
                        }),
                        Metadata {
                            raw: joined,
                            anchor: Some(joined),
                            layer: 0,
                            facts: lm.facts | rm.facts,
                            rule: current_rule,
                            boundary: joined.end,
                        },
                    ),
                    Undo::Binary(lm, rm, ls, rs),
                )
            },
            split_binary,
        );
        changed = true;
    }
    if changed {
        ast.arena.edit(
            left,
            |value| {
                let Value::Ready(expression, meta) = value else {
                    unreachable!()
                };
                let undo = EditUndo::Boundary(meta.boundary, expression.span);
                meta.boundary = raw.end;
                expression.span = span(raw);
                undo
            },
            |value, undo| {
                let (Value::Ready(expression, meta), EditUndo::Boundary(boundary, original)) =
                    (value, undo)
                else {
                    unreachable!()
                };
                meta.boundary = boundary;
                expression.span = original;
            },
        );
        ast.arena.edit(
            left,
            |value| {
                let Value::Ready(_, meta) = value else {
                    unreachable!()
                };
                EditUndo::Anchor(meta.anchor.replace(raw))
            },
            |value, undo| {
                let (Value::Ready(_, meta), EditUndo::Anchor(anchor)) = (value, undo) else {
                    unreachable!()
                };
                meta.anchor = anchor;
            },
        );
    }
    ast.arena.push(left);
}
#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn flattened_comparison_and_logical_order_preserve_exact_spans_and_policy() {
        for source in [
            "const v=a<b<=c===2;",
            "const v=a===2&&b!==3||c??4;",
            "const v=a<<2>3; const w=a>>>2==3;",
            "const v=a|b^c&d;",
            "function f(a,b){if(a<b)return a+b;while(a<4)break;}",
            "const v=a instanceof b;const w=a in b;",
            "const v=a /*comment*/ < b;",
            "const v=a<\n b&&c>2;",
            "const v=typeof x==='string';",
            "const v=f(a<2,b>3);",
            "const v=(a<2)?b:c;const x=a&&b?.x;",
            "const v=document===2||Promise>3;",
        ] {
            let peg = super::super::super::parse(source);
            let owned = super::super::super::parse_direct_bodies(source);
            assert_eq!(format!("{peg:?}"), format!("{owned:?}"), "{source}");
        }
    }
    #[test]
    fn completed_chain_prunes_and_repeatedly_recaptures_only_after_first_move() {
        let source = "document < 2 && Promise > 3";
        let _lines = vybe_ast::line_index::LineIndex::install(source);
        let mut builder = BodyAst::new(source);
        let tree = super::super::super::super::native_grammar::Parser
            .build_captures(Rule::logical_expr, source, &mut builder)
            .unwrap();
        let mut caches = builder.into_caches();
        let cache = caches
            .expressions
            .get_mut(&(0, source.len()))
            .expect("complete logical chain is owned");
        let pair = tree.into_typed_pairs::<Rule>().next().unwrap();
        assert!(pair.into_inner().next().is_none());
        let expected = super::super::super::walk_expression(
            &mut JsWalker::default(),
            JsParser::parse(Rule::logical_expr, source)
                .unwrap()
                .next()
                .unwrap(),
        )
        .unwrap();
        for _ in 0..3 {
            let mut model = false;
            let actual = cache.take_and_record(source, &mut model);
            assert!(model);
            assert_eq!(format!("{actual:?}"), format!("{expected:?}"));
        }
    }
}
