//! Shared .NET AST lowerings.
//!
//! Source languages see their own ByRef/out-param syntax, but the semantics
//! here are .NET surface semantics. Keeping the expression builders here
//! prevents each .NET language frontend from inventing a slightly different
//! TryParse/TryGetValue/ConcurrentCollection rewrite.

use vybe_ast::{Argument, BinOp, ExprKind, Expression, Literal};

use crate::emitter;

fn call_expr(callee: Expression, args: Vec<Argument>) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(callee),
        args,
        optional: false,
    })
}

fn member_expr(object: Expression, field: &str) -> Expression {
    Expression::new(ExprKind::Member {
        object: Box::new(object),
        field: field.to_string(),
        null_safe: false,
    })
}

fn dotted_expr_name(expr: &Expression) -> Option<String> {
    match &expr.kind {
        ExprKind::Ident(name) => Some(name.clone()),
        ExprKind::Member { object, field, .. } => {
            let base = dotted_expr_name(object)?;
            Some(format!("{base}.{field}"))
        }
        _ => None,
    }
}

fn index_expr(object: Expression, index: Expression) -> Expression {
    Expression::new(ExprKind::Index {
        object: Box::new(object),
        index: Box::new(index),
        null_safe: false,
    })
}

fn null_lit() -> Expression {
    Expression::new(ExprKind::Lit(Literal::Null))
}

fn contains_key_expr(object: &Expression, key: &Expression) -> Expression {
    call_expr(
        member_expr(object.clone(), "ContainsKey"),
        vec![Argument::positional(key.clone())],
    )
}

fn assignment_truthy(target: &Expression, value: Expression) -> Expression {
    Expression::new(ExprKind::Binary {
        op: BinOp::Or,
        left: Box::new(Expression::new(ExprKind::Binary {
            op: BinOp::NotEq,
            left: Box::new(Expression::new(ExprKind::Assign {
                target: Box::new(target.clone()),
                value: Box::new(value),
            })),
            right: Box::new(null_lit()),
        })),
        right: Box::new(Expression::bool(true)),
    })
}

pub fn is_hashset_relation_method(field: &str) -> bool {
    matches!(
        field.to_ascii_lowercase().as_str(),
        "issubsetof"
            | "issupersetof"
            | "overlaps"
            | "setequals"
            | "ispropersubsetof"
            | "ispropersupersetof"
    )
}

pub fn collection_base_type_name(type_name: &str) -> String {
    let canonical = emitter::canonical_type_name(type_name);
    canonical
        .split("(Of")
        .next()
        .unwrap_or(&canonical)
        .split('<')
        .next()
        .unwrap_or(&canonical)
        .trim()
        .to_string()
}

pub fn collection_local_type(type_name: &str) -> Option<String> {
    if type_name.contains(".KeyCollection") || type_name.contains(".ValueCollection") {
        return Some("List".into());
    }
    let is_array = type_name.trim().ends_with("()");
    let canonical = collection_base_type_name(type_name);
    let is_collection = emitter::is_component_descriptor_class_in_namespace(
        &canonical,
        "dotnet.System.Collections",
    );
    if !is_collection {
        return None;
    }
    if is_array {
        Some(format!("{canonical}()"))
    } else {
        Some(canonical)
    }
}

pub fn collection_storage_type(type_name: &str) -> &str {
    let type_name = type_name.trim().strip_suffix("()").unwrap_or(type_name);
    if type_name == "DictionaryIgnoreCase" {
        "Dictionary"
    } else {
        type_name
    }
}

pub fn collection_property_method(type_name: &str, field: &str) -> bool {
    emitter::component_instance_method_exists(type_name, field, 0)
}

pub fn collection_type_is_dictionary(type_name: &str) -> bool {
    let storage_type = collection_storage_type(type_name);
    emitter::component_instance_method_exists(storage_type, "ContainsKey", 1)
        && emitter::component_instance_method_exists(storage_type, "Item", 1)
}

pub fn collection_method_takes_dictionary_key(
    type_name: &str,
    field: &str,
    arg_count: usize,
) -> bool {
    let Ok(arg_count) = u8::try_from(arg_count) else {
        return false;
    };
    let storage_type = collection_storage_type(type_name);
    if !emitter::component_instance_method_exists(storage_type, field, arg_count) {
        return false;
    }
    matches!(
        field.to_ascii_lowercase().as_str(),
        "add" | "tryadd" | "containskey" | "item" | "remove" | "getvalueordefault" | "trygetvalue"
    )
}

pub fn datetime_field_name(name: &str) -> &'static str {
    match name.to_ascii_lowercase().as_str() {
        "year" => "Year",
        "month" => "Month",
        "day" => "Day",
        "date" => "Date",
        "hour" => "Hour",
        "minute" => "Minute",
        "second" => "Second",
        "millisecond" => "Millisecond",
        "dayofyear" => "DayOfYear",
        "dayofweek" => "DayOfWeek",
        "ticks" => "Ticks",
        "kind" => "Kind",
        _ => "Year",
    }
}

pub fn is_datetime_static_producer(name: &str) -> bool {
    matches!(
        name.to_ascii_lowercase().as_str(),
        "now"
            | "date"
            | "today"
            | "time"
            | "timeofday"
            | "datevalue"
            | "timevalue"
            | "cdate"
            | "datetime.now"
            | "datetime.utcnow"
            | "datetime.today"
            | "system.datetime.now"
            | "system.datetime.utcnow"
            | "system.datetime.today"
            | "system.datetime.parse"
            | "datetime.minvalue"
            | "datetime.maxvalue"
            | "system.datetime.minvalue"
            | "system.datetime.maxvalue"
    )
}

pub fn encoding_static_name(expr: &Expression) -> Option<&'static str> {
    let receiver = match &expr.kind {
        ExprKind::Call { callee, args, .. } if args.is_empty() => callee.as_ref(),
        _ => expr,
    };
    let path = dotted_expr_name(receiver)?;
    let name = path
        .strip_prefix("Encoding.")
        .or_else(|| path.strip_prefix("System.Text.Encoding."))?;
    match name.to_ascii_lowercase().as_str() {
        "utf8" | "default" => Some("utf8"),
        "ascii" => Some("ascii"),
        "unicode" => Some("unicode"),
        "utf32" => Some("utf32"),
        "latin1" => Some("latin1"),
        "bigendianunicode" => Some("bigendianunicode"),
        _ => None,
    }
}

/// Build a runtime expression that yields the .NET short type name of `expr`.
///
/// Numeric values are represented as f64 in Vybe, so whole numbers map to
/// `Int32` and fractional numbers map to `Double`. Object instances use the
/// shared class stamp when present.
pub fn runtime_type_name_expr(expr: Expression) -> Expression {
    let typeof_expr = Expression::new(ExprKind::TypeOf(Box::new(expr.clone())));

    let type_eq = |runtime_name: &str| {
        Expression::new(ExprKind::Binary {
            op: BinOp::Eq,
            left: Box::new(typeof_expr.clone()),
            right: Box::new(Expression::new(ExprKind::Lit(Literal::Str(
                runtime_name.into(),
            )))),
        })
    };

    let floor_call = Expression::new(ExprKind::Call {
        callee: Box::new(member_expr(Expression::ident("Math"), "floor")),
        args: vec![Argument::positional(expr.clone())],
        optional: false,
    });
    let is_int = Expression::new(ExprKind::Binary {
        op: BinOp::Eq,
        left: Box::new(floor_call),
        right: Box::new(expr.clone()),
    });

    let number_branch = Expression::new(ExprKind::Ternary {
        cond: Box::new(is_int),
        then: Box::new(Expression::string("Int32")),
        else_: Box::new(Expression::string("Double")),
    });

    let inst_type = member_expr(expr, "__type");
    let object_name = Expression::new(ExprKind::Ternary {
        cond: Box::new(inst_type.clone()),
        then: Box::new(inst_type),
        else_: Box::new(Expression::string("Object")),
    });

    let bool_branch = Expression::new(ExprKind::Ternary {
        cond: Box::new(type_eq("boolean")),
        then: Box::new(Expression::string("Boolean")),
        else_: Box::new(object_name),
    });

    let num_or_bool = Expression::new(ExprKind::Ternary {
        cond: Box::new(type_eq("number")),
        then: Box::new(number_branch),
        else_: Box::new(bool_branch),
    });

    Expression::new(ExprKind::Ternary {
        cond: Box::new(type_eq("string")),
        then: Box::new(Expression::string("String")),
        else_: Box::new(num_or_bool),
    })
}

/// `X.TryParse(s, r)` out-param normalization.
pub fn try_parse_desugar(
    recv: Option<&str>,
    callee: &Expression,
    input: &Expression,
    out_target: &Expression,
) -> Option<Expression> {
    let recv = recv?;
    let core = call_expr(callee.clone(), vec![Argument::positional(input.clone())]);
    let assign_core = Expression::new(ExprKind::Assign {
        target: Box::new(out_target.clone()),
        value: Box::new(core.clone()),
    });

    if recv.eq_ignore_ascii_case("Guid")
        || recv.eq_ignore_ascii_case("System.Guid")
        || recv.eq_ignore_ascii_case("Version")
        || recv.eq_ignore_ascii_case("System.Version")
        || recv.eq_ignore_ascii_case("DateTime")
        || recv.eq_ignore_ascii_case("System.DateTime")
        || recv.eq_ignore_ascii_case("Date")
        || recv.eq_ignore_ascii_case("DateTimeOffset")
        || recv.eq_ignore_ascii_case("System.DateTimeOffset")
    {
        let success = Expression::new(ExprKind::Binary {
            op: BinOp::NotEq,
            left: Box::new(core.clone()),
            right: Box::new(null_lit()),
        });
        let assign_success = Expression::new(ExprKind::Binary {
            op: BinOp::NotEq,
            left: Box::new(assign_core),
            right: Box::new(null_lit()),
        });
        return Some(Expression::new(ExprKind::Ternary {
            cond: Box::new(success),
            then: Box::new(assign_success),
            else_: Box::new(Expression::bool(false)),
        }));
    }
    if recv.eq_ignore_ascii_case("Integer")
        || recv.eq_ignore_ascii_case("int")
        || recv.eq_ignore_ascii_case("Int32")
        || recv.eq_ignore_ascii_case("System.Int32")
    {
        let success = Expression::new(ExprKind::Binary {
            op: BinOp::NotEq,
            left: Box::new(assign_core),
            right: Box::new(null_lit()),
        });
        let fallback = Expression::new(ExprKind::Binary {
            op: BinOp::Eq,
            left: Box::new(Expression::new(ExprKind::Assign {
                target: Box::new(out_target.clone()),
                value: Box::new(Expression::new(ExprKind::Lit(Literal::Int(0)))),
            })),
            right: Box::new(null_lit()),
        });
        return Some(Expression::new(ExprKind::Binary {
            op: BinOp::Or,
            left: Box::new(success),
            right: Box::new(fallback),
        }));
    }
    None
}

/// `Uri.TryCreate(input, kind, result)` out-param normalization.
pub fn try_create_desugar(
    recv: Option<&str>,
    callee: &Expression,
    input: &Expression,
    kind: &Expression,
    out_target: &Expression,
) -> Option<Expression> {
    let recv = recv?;
    if !(recv.eq_ignore_ascii_case("Uri") || recv.eq_ignore_ascii_case("System.Uri")) {
        return None;
    }
    let core = call_expr(
        callee.clone(),
        vec![
            Argument::positional(input.clone()),
            Argument::positional(kind.clone()),
        ],
    );
    let assign_core = Expression::new(ExprKind::Assign {
        target: Box::new(out_target.clone()),
        value: Box::new(core),
    });
    Some(Expression::new(ExprKind::Binary {
        op: BinOp::NotEq,
        left: Box::new(assign_core),
        right: Box::new(null_lit()),
    }))
}

/// `Convert.TryFromBase64Chars(chars, dest, bytesWritten)` out-param
/// normalization. The hidden two-arg core returns `[ok, bytesWritten]`.
pub fn try_from_base64_chars_desugar(
    recv: Option<&str>,
    source: &Expression,
    dest: &Expression,
    bytes_written_target: &Expression,
) -> Option<Expression> {
    let recv = recv?;
    if !(recv.eq_ignore_ascii_case("Convert") || recv.eq_ignore_ascii_case("System.Convert")) {
        return None;
    }

    let pair = Expression::ident("__vybe_base64_try_pair");
    let core = call_expr(
        member_expr(
            Expression::new(ExprKind::Ident(recv.to_string())),
            "__TryFromBase64CharsCore",
        ),
        vec![
            Argument::positional(source.clone()),
            Argument::positional(dest.clone()),
        ],
    );
    Some(Expression::new(ExprKind::Sequence(vec![
        Expression::new(ExprKind::Assign {
            target: Box::new(pair.clone()),
            value: Box::new(core),
        }),
        Expression::new(ExprKind::Assign {
            target: Box::new(bytes_written_target.clone()),
            value: Box::new(index_expr(pair.clone(), Expression::int(1))),
        }),
        index_expr(pair, Expression::int(0)),
    ])))
}

/// `d.TryGetValue(k, v)` out-param normalization.
pub fn try_get_value_desugar(
    object: &Expression,
    key: &Expression,
    out_target: &Expression,
) -> Expression {
    try_get_value_desugar_with_default(object, key, out_target, Expression::int(0))
}

pub fn try_get_value_desugar_with_default(
    object: &Expression,
    key: &Expression,
    out_target: &Expression,
    default_value: Expression,
) -> Expression {
    let then_branch = Expression::new(ExprKind::Binary {
        op: BinOp::Or,
        left: Box::new(Expression::new(ExprKind::Binary {
            op: BinOp::NotEq,
            left: Box::new(Expression::new(ExprKind::Assign {
                target: Box::new(out_target.clone()),
                value: Box::new(index_expr(object.clone(), key.clone())),
            })),
            right: Box::new(null_lit()),
        })),
        right: Box::new(Expression::new(ExprKind::Lit(Literal::Bool(true)))),
    });
    let else_branch = Expression::new(ExprKind::Binary {
        op: BinOp::And,
        left: Box::new(Expression::new(ExprKind::Binary {
            op: BinOp::NotEq,
            left: Box::new(Expression::new(ExprKind::Assign {
                target: Box::new(out_target.clone()),
                value: Box::new(default_value),
            })),
            right: Box::new(null_lit()),
        })),
        right: Box::new(Expression::new(ExprKind::Lit(Literal::Bool(false)))),
    });
    Expression::new(ExprKind::Ternary {
        cond: Box::new(contains_key_expr(object, key)),
        then: Box::new(then_branch),
        else_: Box::new(else_branch),
    })
}

pub fn get_or_add_desugar(
    object: &Expression,
    key: &Expression,
    value_factory: &Expression,
) -> Expression {
    let produced = call_expr(
        value_factory.clone(),
        vec![Argument::positional(key.clone())],
    );
    Expression::new(ExprKind::Ternary {
        cond: Box::new(contains_key_expr(object, key)),
        then: Box::new(index_expr(object.clone(), key.clone())),
        else_: Box::new(Expression::new(ExprKind::Assign {
            target: Box::new(index_expr(object.clone(), key.clone())),
            value: Box::new(produced),
        })),
    })
}

pub fn add_or_update_desugar(
    object: &Expression,
    key: &Expression,
    add_value: &Expression,
    update_factory: &Expression,
) -> Expression {
    let current = index_expr(object.clone(), key.clone());
    let updated = call_expr(
        update_factory.clone(),
        vec![
            Argument::positional(key.clone()),
            Argument::positional(current),
        ],
    );
    let target = || index_expr(object.clone(), key.clone());
    Expression::new(ExprKind::Ternary {
        cond: Box::new(contains_key_expr(object, key)),
        then: Box::new(Expression::new(ExprKind::Assign {
            target: Box::new(target()),
            value: Box::new(updated),
        })),
        else_: Box::new(Expression::new(ExprKind::Assign {
            target: Box::new(target()),
            value: Box::new(add_value.clone()),
        })),
    })
}

pub fn try_update_desugar(
    object: &Expression,
    key: &Expression,
    new_value: &Expression,
    comparison_value: &Expression,
) -> Expression {
    let cond = Expression::new(ExprKind::Binary {
        op: BinOp::And,
        left: Box::new(contains_key_expr(object, key)),
        right: Box::new(Expression::new(ExprKind::Binary {
            op: BinOp::Eq,
            left: Box::new(index_expr(object.clone(), key.clone())),
            right: Box::new(comparison_value.clone()),
        })),
    });
    Expression::new(ExprKind::Ternary {
        cond: Box::new(cond),
        then: Box::new(assignment_truthy(
            &index_expr(object.clone(), key.clone()),
            new_value.clone(),
        )),
        else_: Box::new(Expression::bool(false)),
    })
}

pub fn try_remove_desugar(
    object: &Expression,
    key: &Expression,
    out_target: &Expression,
) -> Expression {
    let assign_out = assignment_truthy(out_target, index_expr(object.clone(), key.clone()));
    let remove_call = call_expr(
        member_expr(object.clone(), "Remove"),
        vec![Argument::positional(key.clone())],
    );
    Expression::new(ExprKind::Ternary {
        cond: Box::new(contains_key_expr(object, key)),
        then: Box::new(Expression::new(ExprKind::Binary {
            op: BinOp::And,
            left: Box::new(assign_out),
            right: Box::new(remove_call),
        })),
        else_: Box::new(Expression::bool(false)),
    })
}

pub fn try_take_desugar(object: &Expression, method: &str, out_target: &Expression) -> Expression {
    Expression::new(ExprKind::Ternary {
        cond: Box::new(Expression::new(ExprKind::Binary {
            op: BinOp::Gt,
            left: Box::new(member_expr(object.clone(), "Count")),
            right: Box::new(Expression::int(0)),
        })),
        then: Box::new(assignment_truthy(
            out_target,
            call_expr(
                member_expr(object.clone(), method),
                vec![Argument::positional(out_target.clone())],
            ),
        )),
        else_: Box::new(Expression::bool(false)),
    })
}
