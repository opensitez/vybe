//! Go interface values and method values. Identity is carried by native class
//! descriptors, never by a guest-visible type-name property.
use super::*;
use std::collections::BTreeSet;

const PAYLOAD: &str = "__go_interface_payload";

#[derive(Clone, Default)]
pub(super) struct DeclaredInterfaces {
    pub methods: HashMap<String, HashMap<String, MethodSignature>>,
    pub concrete_types: BTreeSet<String>,
    pub package: String,
    pub source: SourceSignatures,
}

#[derive(Clone, Default)]
pub(super) struct SourceSignatures {
    pub functions: HashMap<String, MethodSignature>,
    pub methods: HashMap<(String, String), MethodSignature>,
}

pub(super) fn source_signatures(source: &str) -> Result<SourceSignatures, String> {
    fn visit(pair: Pair<Rule>, out: &mut SourceSignatures) -> Result<(), String> {
        if matches!(pair.as_rule(), Rule::function_declaration | Rule::method_declaration) {
            let mut name = None;
            let mut receiver = None;
            let mut signature = None;
            for child in pair.into_inner() {
                match child.as_rule() {
                    Rule::ident_name => name = Some(child.as_str().to_owned()),
                    Rule::receiver => {
                        let text = child.as_str().trim().trim_start_matches('(').trim_end_matches(')').trim();
                        let mut parsed = GoParser::parse(Rule::parameter_list, text).map_err(|error| error.to_string())?;
                        if let Some(parameters) = parsed.next() {
                            receiver = walk_parameter_list(parameters)?.first().and_then(|param| param.type_hint.as_deref())
                                .and_then(go_struct_lookup_name);
                        }
                    }
                    Rule::signature => {
                        let source = child.as_str().to_owned();
                        let parsed = walk_signature(child)?;
                        signature = Some(MethodSignature { params: parsed.params,
                            results: results(Some(&source), parsed.return_type.as_deref()), return_type: parsed.return_type });
                    }
                    _ => {}
                }
            }
            if let (Some(name), Some(signature)) = (name, signature) {
                if let Some(receiver) = receiver { out.methods.insert((receiver, name), signature); }
                else { out.functions.insert(name, signature); }
            }
        } else {
            for child in pair.into_inner() { visit(child, out)?; }
        }
        Ok(())
    }
    let source = insert_go_semicolons(source);
    let pairs = GoParser::parse(Rule::program, &source).map_err(|error| error.to_string())?;
    let mut out = SourceSignatures::default();
    for pair in pairs { visit(pair, &mut out)?; }
    Ok(out)
}

#[derive(Clone)]
pub(super) struct MethodSignature {
    pub params: Vec<Param>,
    pub results: Vec<Param>,
    pub return_type: Option<String>,
}

fn parameter(name: &str, ty: Option<&str>) -> Param {
    Param {
        name: name.to_owned(), type_hint: ty.map(Into::into), default: None,
        pass_by: PassBy::Value, is_rest: false, is_kwargs: false,
        is_optional: false, is_nullable: false,
    }
}

fn results(source: Option<&str>, fallback: Option<&str>) -> Vec<Param> {
    if let Some(source) = source {
        if let Ok(mut pairs) = GoParser::parse(Rule::signature, source) {
            if let Some(signature) = pairs.next() {
                for result in signature.into_inner().filter(|p| p.as_rule() == Rule::result) {
                    for part in result.into_inner() {
                        match part.as_rule() {
                            Rule::parameter_list => {
                                if let Ok(params) = walk_parameter_list(part) { return params; }
                            }
                            Rule::type_annotation => return vec![parameter("", Some(&walk_type(part)))],
                            _ => {}
                        }
                    }
                }
            }
        }
    }
    fallback.map(|ty| vec![parameter("", Some(ty))]).unwrap_or_default()
}

pub(super) fn collect(body: &[Statement]) -> DeclaredInterfaces {
    let mut out = DeclaredInterfaces::default();
    let mut parents = HashMap::new();
    for ty in ["bool", "string", "int", "int8", "int16", "int32", "int64", "uint",
        "uint8", "uint16", "uint32", "uint64", "uintptr", "float32", "float64",
        "complex64", "complex128"] {
        out.concrete_types.insert(ty.into());
    }
    for statement in body {
        statement.walk_exprs(&mut |expr| {
            if let ExprKind::Cast { type_name, .. } = &expr.kind {
                if !type_name.is_empty() { out.concrete_types.insert(type_name.clone()); }
            }
        });
        if let StmtKind::StructDecl { name, .. } = &statement.kind {
            out.concrete_types.insert(name.clone());
            out.concrete_types.insert(format!("*{name}"));
        }
        if let StmtKind::VarDecl { declarations, .. } = &statement.kind {
            for declaration in declarations {
                if let Some(ty) = declaration.type_hint.as_deref() { out.concrete_types.insert(ty.to_owned()); }
            }
        }
        if let Some((name, _)) = go_extract_named_type_marker(statement) {
            out.concrete_types.insert(name);
        }
        if let StmtKind::InterfaceDecl { name, parents: inherited, members, .. } = &statement.kind {
            parents.insert(name.clone(), inherited.clone());
            let methods = out.methods.entry(name.clone()).or_default();
            for member in members {
                if let InterfaceMember::Method { name, params, return_type, signature_source, .. } = member {
                    methods.insert(name.clone(), MethodSignature {
                        params: params.clone(),
                        results: results(signature_source.as_deref(), return_type.as_deref()),
                        return_type: return_type.clone(),
                    });
                }
            }
        }
    }
    out.methods.entry("error".into()).or_default().insert("Error".into(), MethodSignature {
        params: Vec::new(), results: vec![parameter("", Some("string"))],
        return_type: Some("string".into()),
    });
    // Fixed point handles forward declarations and arbitrarily deep embedding.
    for _ in 0..parents.len() {
        let previous = out.methods.clone();
        for (name, bases) in &parents {
            let methods = out.methods.entry(name.clone()).or_default();
            for base in bases {
                if let Some(inherited) = previous.get(base) {
                    for (method, signature) in inherited {
                        methods.entry(method.clone()).or_insert_with(|| signature.clone());
                    }
                }
            }
        }
    }
    out.concrete_types.retain(|ty| !out.methods.contains_key(ty)
        && !matches!(ty.as_str(), "any" | "interface{}"));
    out
}

pub(super) fn method<'a>(env: &'a GoNormalizeEnv, ty: &str, name: &str) -> Option<&'a MethodSignature> {
    env.interface_signatures.methods.get(ty.trim())?.get(name)
}

pub(super) fn concrete_method(env: &GoNormalizeEnv, ty: &str, name: &str) -> Option<MethodSignature> {
    fn find(env: &GoNormalizeEnv, ty: &str, name: &str, seen: &mut HashSet<String>) -> Option<MethodSignature> {
        let lookup = go_struct_lookup_name(ty)?;
        if !seen.insert(lookup.clone()) { return None; }
        let info = env.struct_infos.get(&lookup)?;
        if let Some(params) = info.method_params.get(name) {
            if let Some(signature) = env.interface_signatures.source.methods.get(&(lookup, name.to_owned())) {
                return Some(signature.clone());
            }
            let return_type = info.member_types.get(name).cloned();
            return Some(MethodSignature { params: params.iter().skip(1).cloned().collect(),
                results: results(None, return_type.as_deref()), return_type });
        }
        for (_, embedded) in &info.embedded_fields {
            if let Some(signature) = method(env, embedded, name) { return Some(signature.clone()); }
            if let Some(signature) = find(env, embedded, name, seen) { return Some(signature); }
        }
        None
    }
    find(env, ty, name, &mut HashSet::new())
}

pub(super) fn function_type(signature: &MethodSignature) -> String {
    let params = signature.params.iter().map(|param| {
        let ty = param.type_hint.as_deref().unwrap_or("any");
        if param.is_rest { format!("...{}", ty.strip_prefix("[]").unwrap_or(ty)) } else { ty.to_owned() }
    }).collect::<Vec<_>>().join(",");
    let result = signature.results.iter().map(|param| param.type_hint.as_deref().unwrap_or("any"))
        .collect::<Vec<_>>().join(",");
    if signature.results.is_empty() { format!("func({params})") }
    else { format!("func({params})({result})") }
}

pub(super) fn function_signature(ty: &str) -> Option<MethodSignature> {
    let source = ty.trim().strip_prefix("func")?;
    let signature = GoParser::parse(Rule::signature, source).ok()?.next()?;
    let parsed = walk_signature(signature).ok()?;
    Some(MethodSignature { params: parsed.params, results: results(Some(source), parsed.return_type.as_deref()),
        return_type: parsed.return_type })
}

pub(super) fn member_function_type(expr: &Expression, env: &GoNormalizeEnv,
    functions: &HashMap<String, GoFunctionSignature>) -> Option<String> {
    let ExprKind::Member { object, field, .. } = &expr.kind else { return None; };
    let ty = go_expr_type_hint(object, env, functions)?;
    if let Some(signature) = method(env, &ty, field) { return Some(function_type(signature)); }
    let signature = concrete_method(env, &ty, field)?;
    // The older concrete AST uses [N] for tuple returns; do not invent their
    // element types. Interface declarations retain the full result list.
    if signature.return_type.as_deref().is_some_and(|ty| ty.starts_with('[') && ty.ends_with(']')) { return None; }
    Some(function_type(&signature))
}

pub(super) fn carrier_name(ty: &str) -> String {
    let mut name = String::from("__go_interface_carrier_");
    // This is only a hygienic compiler symbol, not runtime type identity.
    for byte in ty.trim().bytes() { name.push_str(&format!("{byte:02x}")); }
    name
}

fn member(object: Expression, name: &str) -> Expression {
    Expression::new(ExprKind::Member { object: Box::new(object), field: name.into(), null_safe: false })
}

fn is_null(value: Expression) -> Expression {
    Expression::new(ExprKind::Binary { op: BinOp::StrictEq,
        left: Box::new(value), right: Box::new(Expression::null()) })
}

fn lambda(params: Vec<Param>, body: LambdaBody, captures: Vec<String>) -> Expression {
    Expression::new(ExprKind::Lambda { params, body, captures, is_async: false })
}

pub(super) fn with_snapshot(
    value: Expression, name: String, ty: Option<&str>, body: Vec<Statement>,
) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(lambda(vec![parameter(&name, ty)], LambdaBody::Block(body), Vec::new())),
        args: vec![Argument::positional(value)], optional: false,
    })
}

pub(super) fn require_non_nil(value: Expression, state: &mut GoNormalizeState) -> Expression {
    let name = fresh_go_temp(state, "__go_deref_receiver");
    with_snapshot(value, name.clone(), None, vec![
        Statement::new(StmtKind::If { cond: is_null(Expression::ident(&name)),
            then_body: vec![Statement::new(StmtKind::Throw {
                expr: Some(Expression::string("invalid memory address or nil pointer dereference")), cause: None,
            })], elifs: Vec::new(), else_body: None }),
        Statement::new(StmtKind::Return(Some(Expression::ident(&name)))),
    ])
}

/// A nil function panics on invocation, AFTER its argument expressions have
/// been evaluated. Keep the callee and each argument once; a spread argument
/// remains the original slice value, not a flattened/copied snapshot.
pub(super) fn guard_function_call(
    callee: Expression, ty: &str, args: Vec<Argument>, optional: bool,
    state: &mut GoNormalizeState,
) -> Expression {
    let snapshot = fresh_go_temp(state, "__go_function_value");
    let mut body = Vec::new();
    let args = args.into_iter().map(|mut arg| {
        let name = fresh_go_temp(state, "__go_function_arg");
        body.push(go_defer_temp_decl(name.clone(), None, arg.value));
        arg.value = Expression::ident(&name);
        arg
    }).collect();
    body.push(Statement::new(StmtKind::If {
        cond: is_null(Expression::ident(&snapshot)),
        then_body: vec![Statement::new(StmtKind::Throw {
            expr: Some(Expression::string("invalid memory address or nil pointer dereference")),
            cause: None,
        })], elifs: Vec::new(), else_body: None,
    }));
    body.push(Statement::new(StmtKind::Return(Some(Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident(&snapshot)), args, optional,
    })))));
    with_snapshot(callee, snapshot, Some(ty), body)
}

pub(super) fn promoted_receiver(mut object: Expression, ty: &str, name: &str,
    env: &GoNormalizeEnv, state: &mut GoNormalizeState) -> Option<(Expression, String)> {
    fn path(ty: &str, name: &str, env: &GoNormalizeEnv,
        seen: &mut HashSet<String>) -> Option<Vec<(String, String)>> {
        let lookup = go_struct_lookup_name(ty)?;
        if !seen.insert(lookup.clone()) { return None; }
        let info = env.struct_infos.get(&lookup)?;
        if info.method_names.contains(name) || method(env, ty, name).is_some() { return Some(Vec::new()); }
        for (field, embedded) in &info.embedded_fields {
            if method(env, embedded, name).is_some() { return Some(vec![(field.clone(), embedded.clone())]); }
            if let Some(mut tail) = path(embedded, name, env, &mut seen.clone()) {
                tail.insert(0, (field.clone(), embedded.clone())); return Some(tail);
            }
        }
        None
    }
    let path = path(ty, name, env, &mut HashSet::new())?;
    let mut current_type = ty.to_owned();
    for (field, embedded_type) in path {
        object = if current_type.starts_with('*') {
            Expression::new(ExprKind::RefLoad(Box::new(require_non_nil(object, state))))
        } else if let ExprKind::RefLoad(pointer) = object.kind {
            Expression::new(ExprKind::RefLoad(Box::new(require_non_nil(*pointer, state))))
        } else { object };
        object = member(object, &field);
        current_type = embedded_type;
    }
    Some((object, current_type))
}

pub(super) fn tuple_hints(expr: &Expression, env: &GoNormalizeEnv,
    functions: &HashMap<String, GoFunctionSignature>) -> Option<Vec<Option<String>>> {
    let ExprKind::Call { callee, args, .. } = &expr.kind else { return None; };
    if matches!(&callee.kind, ExprKind::Ident(name) if name == "__go_type_assert_comma_ok") {
        return Some(vec![args.get(1).and_then(|arg| go_type_name_from_expr(&arg.value)), Some("bool".into())]);
    }
    let signature = if let ExprKind::Ident(name) = &callee.kind {
        env.interface_signatures.source.functions.get(name).cloned()
            .or_else(|| env.value_types.get(name).and_then(|ty| function_signature(ty)))
    } else if let ExprKind::Member { object, field, .. } = &callee.kind {
        let ty = go_expr_type_hint(object, env, functions)?;
        method(env, &ty, field).cloned().or_else(|| concrete_method(env, &ty, field))
    } else { go_expr_type_hint(callee, env, functions).and_then(|ty| function_signature(&ty)) }?;
    Some(signature.results.iter().map(|result| result.type_hint.as_deref().map(str::to_owned)).collect())
}

pub(super) fn call_parameter_type(callee: &Expression, index: usize, spread: bool,
    env: &GoNormalizeEnv, functions: &HashMap<String, GoFunctionSignature>) -> Option<String> {
    let signature = if let ExprKind::Ident(name) = &callee.kind {
        env.interface_signatures.source.functions.get(name).cloned()
            .or_else(|| env.value_types.get(name).and_then(|ty| function_signature(ty)))
    } else if let ExprKind::Member { object, field, .. } = &callee.kind {
        let ty = go_expr_type_hint(object, env, functions)?;
        method(env, &ty, field).cloned().or_else(|| concrete_method(env, &ty, field))
    } else { go_expr_type_hint(callee, env, functions).and_then(|ty| function_signature(&ty)) }?;
    let param = signature.params.get(index).or_else(|| signature.params.last().filter(|param| param.is_rest))?;
    let ty = param.type_hint.as_deref()?;
    Some(if param.is_rest && !spread { ty.strip_prefix("[]").unwrap_or(ty).to_owned() } else { ty.to_owned() })
}

pub(super) fn declaration_signature(name: &str, params: &[Param], env: &GoNormalizeEnv) -> Option<MethodSignature> {
    env.interface_signatures.source.functions.get(name).cloned().or_else(|| {
        let receiver = params.first()?.type_hint.as_deref().and_then(go_struct_lookup_name)?;
        env.interface_signatures.source.methods.get(&(receiver, name.to_owned())).cloned()
    })
}

pub(super) fn return_value(mut value: Expression, source: &Expression, env: &GoNormalizeEnv,
    functions: &HashMap<String, GoFunctionSignature>, state: &mut GoNormalizeState) -> Expression {
    if let Some(signature) = &env.return_signature {
        if signature.results.len() > 1 {
            if let (ExprKind::Array(values), ExprKind::Array(sources)) = (&mut value.kind, &source.kind) {
                for ((value, source), result) in values.iter_mut().zip(sources).zip(&signature.results) {
                    if let Some(target) = result.type_hint.as_deref() {
                        value.value = box_value(value.value.clone(), &source.value, target, env, functions, state);
                    }
                }
                return value;
            }
            let values = match &mut value.kind { ExprKind::Tuple(values) => Some(values), _ => None };
            let sources = match &source.kind { ExprKind::Tuple(values) => Some(values), _ => None };
            if let (Some(values), Some(sources)) = (values, sources) {
                for ((value, source), result) in values.iter_mut().zip(sources).zip(&signature.results) {
                    if let Some(target) = result.type_hint.as_deref() {
                        *value = box_value(value.clone(), source, target, env, functions, state);
                    }
                }
                return value;
            }
        }
    }
    if let Some(target) = env.return_type.as_deref() { box_value(value, source, target, env, functions, state) }
    else { value }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn nil_function_guard_evaluates_arguments_before_testing_callee() {
        let mut state = GoNormalizeState::default();
        let expression = guard_function_call(Expression::ident("callback"), "func(...int32) int32",
            vec![Argument { value: go_builtin_call("argumentEffect", Vec::new()),
                name: None, by_ref: false, spread: true }], false, &mut state);
        let ExprKind::Call { callee, args, .. } = expression.kind else { panic!("callee snapshot"); };
        assert_eq!(args.len(), 1);
        assert!(matches!(&args[0].value.kind, ExprKind::Ident(name) if name == "callback"));
        let ExprKind::Lambda { body: LambdaBody::Block(body), .. } = callee.kind else { panic!("guard body"); };
        assert!(matches!(body[0].kind, StmtKind::VarDecl { .. }));
        assert!(matches!(body[1].kind, StmtKind::If { .. }));
        let StmtKind::Return(Some(Expression { kind: ExprKind::Call { args, .. }, .. })) = &body[2].kind
            else { panic!("actual invocation"); };
        assert!(args[0].spread);
        assert!(matches!(args[0].value.kind, ExprKind::Ident(_)));
    }

    #[test]
    fn named_result_defer_factory_forwards_recovery_callable() {
        let env = GoNormalizeEnv { recover_fn_name: Some("recovery".into()), ..Default::default() };
        let body = LambdaBody::Expr(Box::new(Expression::new(ExprKind::Sequence(vec![
            Expression::new(ExprKind::RefOf(Box::new(PlaceExpr::Ident("result".into())))),
            go_builtin_call("recovery", Vec::new()),
        ]))));
        let expression = go_deferred_lambda_with_ref_captures(body, &[], &env, &HashMap::new(),
            &mut GoNormalizeState::default());
        let ExprKind::Call { callee, args, .. } = expression.kind else { panic!("capture factory"); };
        assert!(args.iter().any(|arg| matches!(&arg.value.kind, ExprKind::Ident(name) if name == "recovery")));
        let ExprKind::Lambda { params, .. } = callee.kind else { panic!("factory parameters"); };
        assert!(params.iter().any(|param| param.name == "recovery"));
    }

    #[test]
    fn reflection_singletons_are_emitted_only_when_demanded() {
        let mut env = GoNormalizeEnv::default();
        env.interface_signatures = collect(&[]);
        let mut state = GoNormalizeState::default();
        let declarations = carrier_declarations(&env, &state);
        assert!(!declarations.iter().any(|stmt| matches!(stmt.kind, StmtKind::VarDecl { .. })));
        let source = Expression::new(ExprKind::Lit(Literal::Int(7)));
        let _ = reflect(source.clone(), &source, false, &env, &HashMap::new(), &mut state);
        assert_eq!(state.interface_reflection_types, BTreeSet::from(["int".to_owned()]));
        let declarations = carrier_declarations(&env, &state);
        assert_eq!(declarations.iter().filter(|stmt| matches!(stmt.kind, StmtKind::VarDecl { .. })).count(), 1);
    }

    #[test]
    fn assertion_snapshot_argument_is_the_original_effectful_operand() {
        let operand = go_builtin_call("receiverFactory", Vec::new());
        let result = assertion(operand, "int32", &GoNormalizeEnv::default(),
            &mut GoNormalizeState::default(), true);
        let ExprKind::Call { args, callee, .. } = result.kind else { panic!("snapshot factory call"); };
        assert_eq!(args.len(), 1);
        let ExprKind::Call { callee: original, .. } = &args[0].value.kind else { panic!("original operand retained"); };
        assert!(matches!(&original.kind, ExprKind::Ident(name) if name == "receiverFactory"));
        assert!(matches!(callee.kind, ExprKind::Lambda { .. }));
    }

    #[test]
    fn anonymous_interface_conversion_retains_its_type() {
        let pair = GoParser::parse(Rule::interface_conversion, "interface{}(value)")
            .unwrap().next().unwrap();
        let expression = walk_type_conversion(pair).unwrap();
        assert!(matches!(expression.kind, ExprKind::Cast { type_name, .. } if type_name == "interface{}"));
    }

    #[test]
    fn method_results_and_variadic_parameters_are_retained_from_source() {
        let source = "package main\ntype Box struct{}\nfunc (b *Box) Values(values ...int32) (int32, bool) { return 1, true }\nfunc main() {}";
        let metadata = source_signatures(source).unwrap();
        let signature = metadata.methods.get(&("Box".into(), "Values".into())).unwrap();
        assert_eq!(signature.params.len(), 1);
        assert!(signature.params[0].is_rest);
        assert_eq!(signature.results.len(), 2);
        assert_eq!(signature.results[0].type_hint.as_deref(), Some("int32"));
        assert_eq!(signature.results[1].type_hint.as_deref(), Some("bool"));
    }

    #[test]
    fn interface_binding_returns_an_actual_lambda_after_its_nil_check() {
        let mut env = GoNormalizeEnv::default();
        env.interface_signatures.methods.insert("Stepper".into(), HashMap::from([("Step".into(),
            MethodSignature { params: Vec::new(), results: vec![parameter("", Some("int32"))],
                return_type: Some("int32".into()) })]));
        let expression = materialize(go_builtin_call("receiverFactory", Vec::new()), "Stepper", "Step",
            &env, &mut GoNormalizeState::default()).unwrap();
        let ExprKind::Call { callee, args, .. } = expression.kind else { panic!("binding factory"); };
        assert_eq!(args.len(), 1);
        let ExprKind::Lambda { body: LambdaBody::Block(body), .. } = callee.kind else { panic!("ordered binding body"); };
        assert!(matches!(body[0].kind, StmtKind::If { .. }));
        assert!(matches!(&body[1].kind, StmtKind::Return(Some(Expression { kind: ExprKind::Lambda { .. }, .. }))));
    }
}

pub(super) fn composite(mut value: Expression, source: &Expression, target: &str,
    env: &GoNormalizeEnv, functions: &HashMap<String, GoFunctionSignature>, state: &mut GoNormalizeState) -> Expression {
    let element_type = go_array_element_type(target).or_else(|| go_map_value_type(target));
    match (&mut value.kind, &source.kind) {
        (ExprKind::Array(values), ExprKind::Array(sources)) => {
            if let Some(element_type) = element_type.as_deref() {
                for (value, source) in values.iter_mut().zip(sources) {
                    value.value = box_value(value.value.clone(), &source.value, element_type, env, functions, state);
                }
            }
        }
        (ExprKind::Object(values), ExprKind::Object(sources)) => {
            let info = go_struct_lookup_name(target).and_then(|name| env.struct_infos.get(&name));
            for (value, source) in values.iter_mut().zip(sources) {
                if let (ObjectProperty::KeyValue { key, value }, ObjectProperty::KeyValue { value: original, .. }) = (value, source) {
                    let field_type = if let ExprKind::Lit(Literal::Str(name)) = &key.kind {
                        info.and_then(|info| info.member_types.get(name)).map(String::as_str)
                    } else { None };
                    if let Some(ty) = field_type.or(element_type.as_deref()) {
                        *value = box_value(value.clone(), original, ty, env, functions, state);
                    }
                }
            }
        }
        _ => {}
    }
    value
}

pub(super) fn box_value(
    value: Expression, source: &Expression, target: &str, env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>, state: &mut GoNormalizeState,
) -> Expression {
    if !go_is_go_interface_type(target, env) || go_is_null_expr(source) { return value; }
    let Some(ty) = go_expr_type_hint(source, env, signatures) else { return value; };
    if go_is_go_interface_type(&ty, env) || ty.starts_with("__go_interface_carrier_") { return value; }
    state.interface_carriers.insert(ty.clone());
    let value = go_wrap_go_value_copy(value, env, signatures);
    Expression::new(ExprKind::Cast {
        expr: Box::new(Expression::new(ExprKind::Object(vec![ObjectProperty::KeyValue {
            key: Expression::string(PAYLOAD), value,
        }]))), type_name: carrier_name(&ty),
    })
}

pub(super) fn materialize(
    receiver: Expression, receiver_type: &str, name: &str, env: &GoNormalizeEnv,
    state: &mut GoNormalizeState,
) -> Option<Expression> {
    let signature = method(env, receiver_type, name)?;
    let snapshot = fresh_go_temp(state, "__go_bound_interface");
    let mut params = signature.params.clone();
    for (index, param) in params.iter_mut().enumerate() {
        param.name = format!("{snapshot}_arg_{index}");
    }
    let args = params.iter().map(|param| Argument {
        value: Expression::ident(&param.name), name: None, by_ref: false, spread: param.is_rest,
    }).collect();
    let call = Expression::new(ExprKind::Call {
        callee: Box::new(member(Expression::ident(&snapshot), name)), args, optional: false,
    });
    let callable = lambda(params, LambdaBody::Expr(Box::new(call)), vec![snapshot.clone()]);
    Some(with_snapshot(receiver, snapshot.clone(), None, vec![
        Statement::new(StmtKind::If { cond: is_null(Expression::ident(&snapshot)),
            then_body: vec![Statement::new(StmtKind::Throw {
                expr: Some(Expression::string("invalid memory address or nil pointer dereference")), cause: None,
            })], elifs: Vec::new(), else_body: None }),
        Statement::new(StmtKind::Return(Some(callable))),
    ]))
}

fn type_test(value: Expression, ty: &str) -> Expression {
    Expression::new(ExprKind::IsType { expr: Box::new(value), type_name: carrier_name(ty) })
}

pub(super) fn test(value: Expression, target: &str, env: &GoNormalizeEnv) -> Expression {
    if target == "__go_nil" { return is_null(value); }
    if !go_is_go_interface_type(target, env) { return type_test(value, target); }
    env.interface_signatures.concrete_types.iter()
        .filter(|ty| go_type_assignable_to_interface(ty, target, env))
        .map(|ty| type_test(value.clone(), ty))
        .reduce(|left, right| Expression::new(ExprKind::Binary {
            op: BinOp::Or, left: Box::new(left), right: Box::new(right),
        })).unwrap_or_else(|| Expression::bool(false))
}

pub(super) fn equal(left: Expression, right: Expression, op: BinOp,
    env: &GoNormalizeEnv, state: &mut GoNormalizeState) -> Expression {
    let left_name = fresh_go_temp(state, "__go_interface_eq_left");
    let right_name = fresh_go_temp(state, "__go_interface_eq_right");
    let l = Expression::ident(&left_name);
    let r = Expression::ident(&right_name);
    let mut body = vec![Statement::new(StmtKind::If { cond: is_null(l.clone()),
        then_body: vec![Statement::new(StmtKind::Return(Some(is_null(r.clone()))))],
        elifs: Vec::new(), else_body: None })];
    for ty in &env.interface_signatures.concrete_types {
        let mut matched = vec![Statement::new(StmtKind::If { cond: Expression::new(ExprKind::Unary {
            op: UnaryOp::Not, expr: Box::new(type_test(r.clone(), ty)),
        }), then_body: vec![Statement::new(StmtKind::Return(Some(Expression::bool(false))))],
            elifs: Vec::new(), else_body: None })];
        if go_is_map_type(ty) || ty.starts_with("[]") || ty.starts_with("func") {
            matched.push(Statement::new(StmtKind::Throw { expr: Some(Expression::string("comparing uncomparable interface value")), cause: None }));
        } else {
            let lp = Expression::new(ExprKind::Cast { expr: Box::new(member(l.clone(), PAYLOAD)), type_name: ty.clone() });
            let rp = Expression::new(ExprKind::Cast { expr: Box::new(member(r.clone(), PAYLOAD)), type_name: ty.clone() });
            let comparison = if ty.starts_with('*') {
                Expression::new(ExprKind::Binary { op: BinOp::Eq, left: Box::new(lp), right: Box::new(rp) })
            } else {
                go_struct_equality_expr(lp.clone(), rp.clone(), BinOp::Eq, env, &HashMap::new())
                    .unwrap_or_else(|| Expression::new(ExprKind::Binary { op: BinOp::Eq, left: Box::new(lp), right: Box::new(rp) }))
            };
            matched.push(Statement::new(StmtKind::Return(Some(comparison))));
        }
        body.push(Statement::new(StmtKind::If { cond: type_test(l.clone(), ty), then_body: matched,
            elifs: Vec::new(), else_body: None }));
    }
    body.push(Statement::new(StmtKind::Return(Some(Expression::bool(false)))));
    let call = Expression::new(ExprKind::Call {
        callee: Box::new(lambda(vec![parameter(&left_name, None), parameter(&right_name, None)],
            LambdaBody::Block(body), Vec::new())),
        args: vec![Argument::positional(left), Argument::positional(right)], optional: false,
    });
    if op == BinOp::NotEq { Expression::new(ExprKind::Unary { op: UnaryOp::Not, expr: Box::new(call) }) } else { call }
}

fn reflection_name(ty: &str) -> String { format!("{}_reflect_type", carrier_name(ty)) }

fn display_type(ty: &str, env: &GoNormalizeEnv) -> String {
    if let Some(inner) = ty.strip_prefix('*') { return format!("*{}", display_type(inner, env)); }
    if env.struct_infos.contains_key(ty) || env.named_types.contains_key(ty) {
        format!("{}.{ty}", env.interface_signatures.package)
    } else { go_reflect_display_type(ty) }
}

pub(super) fn reflect(value: Expression, source: &Expression, value_of: bool,
    env: &GoNormalizeEnv, functions: &HashMap<String, GoFunctionSignature>, state: &mut GoNormalizeState) -> Expression {
    let source_type = go_expr_type_hint(source, env, functions);
    if let Some(ty) = source_type.as_deref().filter(|ty| !go_is_go_interface_type(ty, env)) {
        if !value_of {
            // TypeOf still evaluates an effectful operand even though its
            // dynamic type is known at this concrete conversion boundary.
            state.interface_reflection_types.insert(ty.to_owned());
            return Expression::new(ExprKind::Sequence(vec![value, Expression::ident(&reflection_name(ty))]));
        }
        return go_builtin_call("__go_reflect_valueof", vec![value,
            Expression::string(&display_type(ty, env)), Expression::string(&go_reflect_kind_name(ty, env))]);
    }
    let snapshot = fresh_go_temp(state, "__go_reflect_interface");
    let mut body = Vec::new();
    for ty in &env.interface_signatures.concrete_types {
        if !value_of { state.interface_reflection_types.insert(ty.clone()); }
        let result = if value_of {
            go_builtin_call("__go_reflect_valueof", vec![member(Expression::ident(&snapshot), PAYLOAD),
                Expression::string(&display_type(ty, env)), Expression::string(&go_reflect_kind_name(ty, env))])
        } else { Expression::ident(&reflection_name(ty)) };
        body.push(Statement::new(StmtKind::If { cond: type_test(Expression::ident(&snapshot), ty),
            then_body: vec![Statement::new(StmtKind::Return(Some(result)))], elifs: Vec::new(), else_body: None }));
    }
    let missing = if value_of { go_builtin_call("__go_reflect_valueof", vec![Expression::null(),
        Expression::string(""), Expression::string("invalid")]) } else { Expression::null() };
    body.push(Statement::new(StmtKind::Return(Some(missing))));
    with_snapshot(value, snapshot, None, body)
}

pub(super) fn assertion(
    value: Expression, target: &str, env: &GoNormalizeEnv,
    state: &mut GoNormalizeState, comma_ok: bool,
) -> Expression {
    let operand = value;
    let snapshot = fresh_go_temp(state, "__go_assert_interface");
    let value = Expression::ident(&snapshot);
    let valid = test(value.clone(), target, env);
    let result = if go_is_go_interface_type(target, env) { value.clone() } else {
        Expression::new(ExprKind::Cast { expr: Box::new(member(value.clone(), PAYLOAD)), type_name: target.into() })
    };
    let body = if comma_ok {
        vec![Statement::new(StmtKind::Return(Some(Expression::new(ExprKind::Tuple(vec![
            Expression::new(ExprKind::Ternary { cond: Box::new(valid.clone()),
                then: Box::new(result), else_: Box::new(go_zero_value_for_type(target, env)) }), valid,
        ])))))]
    } else {
        vec![Statement::new(StmtKind::If { cond: valid,
            then_body: vec![Statement::new(StmtKind::Return(Some(result)))],
            elifs: Vec::new(), else_body: None }),
            Statement::new(StmtKind::Throw { expr: Some(Expression::string("interface conversion failed")), cause: None })]
    };
    with_snapshot(operand, snapshot, None, body)
}

pub(super) fn carrier_declarations(env: &GoNormalizeEnv, state: &GoNormalizeState) -> Vec<Statement> {
    let mut declarations = Vec::new();
    let mut types = env.interface_signatures.concrete_types.clone();
    types.extend(state.interface_carriers.iter().cloned());
    types.extend(state.interface_reflection_types.iter().cloned());
    for ty in &types {
        let name = carrier_name(ty);
        let mut members = vec![ClassMember::Field { name: PAYLOAD.into(),
            type_hint: Some(ty.clone()), init: None, modifiers: Modifiers::default(),
            array_bounds: None, with_events: false, storage: None }];
        let mut emitted = HashSet::new();
        for methods in env.interface_signatures.methods.values() {
            for method in methods.keys() {
                if !go_type_has_method(ty, method, env) || !emitted.insert(method.clone()) { continue; }
                let Some(signature) = concrete_method(env, ty, method) else { continue; };
                let receiver = format!("__go_carrier_self_{method}");
                let payload = format!("__go_carrier_payload_{method}");
                let mut params = signature.params.clone();
                for (index, param) in params.iter_mut().enumerate() { param.name = format!("__go_carrier_arg_{index}"); }
                let args = params.iter().map(|param| Argument { value: Expression::ident(&param.name),
                    name: None, by_ref: false, spread: param.is_rest }).collect::<Vec<_>>();
                let mut local = env.clone();
                local.value_types.insert(payload.clone(), ty.clone());
                // This adapter runs on INVOCATION. Promotion and a value-method
                // wrapper's implicit dereference must not happen at interface
                // method-value binding time.
                let mut receiver_state = GoNormalizeState::default();
                let (mut target, target_type) = promoted_receiver(Expression::ident(&payload), ty, method,
                    &local, &mut receiver_state).unwrap_or_else(|| (Expression::ident(&payload), ty.clone()));
                let needs_value = go_struct_lookup_name(&target_type).and_then(|name| local.struct_infos.get(&name))
                    .is_some_and(|info| info.method_names.contains(method) && !info.pointer_method_names.contains(method));
                if needs_value && target_type.starts_with('*') {
                    target = require_non_nil(target, &mut receiver_state);
                }
                // Promotion can insert a guarded IIFE before a field load.
                // Preserve the declared embedded receiver type across that
                // expression, including pointer methods (not only value ones).
                target = Expression::new(ExprKind::Cast {
                    expr: Box::new(target), type_name: target_type,
                });
                let call = go_bound_method_call_expr(target, method, &args, false, &local, &HashMap::new())
                    .unwrap_or_else(|| Expression::new(ExprKind::Call {
                        callee: Box::new(member(Expression::ident(&payload), method)), args, optional: false,
                    }));
                params.insert(0, parameter(&receiver, Some(&name)));
                members.push(ClassMember::Method(Box::new(Statement::new(StmtKind::FunctionDecl {
                    name: method.clone(), params, return_type: signature.return_type.clone(),
                    body: vec![go_defer_temp_decl(payload, Some(ty.clone()), member(Expression::ident(&receiver), PAYLOAD)),
                        Statement::new(StmtKind::Return(Some(call)))],
                    modifiers: Modifiers::default(), handles: Vec::new(),
                    is_async: false, is_generator: false, is_sub: false,
                }))));
            }
        }
        declarations.push(Statement::new(StmtKind::StructDecl { name, interfaces: Vec::new(), members,
            visibility: Visibility::Private, decorators: Vec::new(), semantics: ValueSemantics::default() }));
        if state.interface_reflection_types.contains(ty) {
            declarations.push(go_defer_temp_decl(reflection_name(ty), Some("__goReflectType".into()),
                go_builtin_call("__go_reflect_typeof", vec![Expression::null(),
                    Expression::string(&display_type(ty, env)), Expression::string(&go_reflect_kind_name(ty, env)),
                    go_reflect_fields_expr(ty, env)])));
        }
    }
    declarations
}
