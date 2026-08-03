//! Shared generics primitive.
//!
//! Generics are not a WASM feature in Vybe. Language walkers normalize generic
//! syntax into `vybe_ast::TypeRef`; this module canonicalizes, substitutes, and
//! prepares metadata so classes, reflection, collections, and overloads can
//! share one generic model.

use std::collections::HashMap;
use vybe_ast::{
    GenericArg, GenericBound, GenericConstraint, GenericParam, GenericRuntimeMode, GenericVariance,
    TypePath, TypeRef, TypeRefKind };

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct CanonicalType {
    pub name: String }

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct GenericKey {
    pub owner_path: String,
    pub args: Vec<CanonicalType>,
    pub mode: GenericRuntimeMode }

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct GenericSignature {
    pub params: Vec<GenericParam> }

impl GenericSignature {
    pub fn new(params: Vec<GenericParam>) -> Self {
        Self { params }
    }

    pub fn bind_args(&self, args: &[TypeRef]) -> Result<GenericContext, GenericError> {
        check_arity(args, &self.params)?;
        let substitutions: Vec<(String, TypeRef)> = self
            .params
            .iter()
            .enumerate()
            .filter_map(|(index, param)| {
                args.get(index)
                    .cloned()
                    .or_else(|| param.default.clone())
                    .map(|arg| (param.name.clone(), arg))
            })
            .collect();
        Ok(GenericContext::from_pairs(substitutions))
    }
}

pub fn runtime_type_arg_param_name(name: &str) -> String {
    format!("__generic_typearg_{}", sanitize_generic_runtime_name(name))
}

pub fn runtime_type_arg_param_names(params: &[GenericParam]) -> Vec<String> {
    params
        .iter()
        .map(|param| runtime_type_arg_param_name(&param.name))
        .collect()
}

fn sanitize_generic_runtime_name(name: &str) -> String {
    let mut out = String::new();
    for ch in name.chars() {
        if ch.is_ascii_alphanumeric() || ch == '_' {
            out.push(ch);
        } else if !out.ends_with('_') {
            out.push('_');
        }
    }
    let trimmed = out.trim_matches('_');
    if trimmed.is_empty() {
        "T".to_string()
    } else {
        trimmed.to_string()
    }
}

#[derive(Debug, Clone, Default)]
pub struct GenericContext {
    substitutions: HashMap<String, TypeRef> }

impl GenericContext {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn with_substitution(mut self, name: impl Into<String>, type_ref: TypeRef) -> Self {
        self.substitutions.insert(name.into(), type_ref);
        self
    }

    pub fn from_pairs(pairs: impl IntoIterator<Item = (String, TypeRef)>) -> Self {
        Self {
            substitutions: pairs.into_iter().collect() }
    }

    pub fn get(&self, name: &str) -> Option<&TypeRef> {
        self.substitutions.get(name)
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum GenericError {
    ArityMismatch { expected: usize, actual: usize },
    ConstraintViolation { param: String, reason: String } }

pub fn canonical_type_ref(type_ref: &TypeRef) -> CanonicalType {
    CanonicalType {
        name: display_type_ref(type_ref) }
}

pub fn canonical_generic_key(
    owner_path: impl Into<String>,
    args: &[TypeRef],
    mode: GenericRuntimeMode,
) -> GenericKey {
    GenericKey {
        owner_path: owner_path.into(),
        args: args.iter().map(canonical_type_ref).collect(),
        mode }
}

pub fn erased_type_name(type_name: &str) -> String {
    parse_type_ref_hint(type_name)
        .map(|ty| display_type_ref(&erase_type(&ty)))
        .unwrap_or_else(|| strip_generic_application_text(type_name).trim().to_string())
}

pub fn generic_base_name(type_name: &str) -> &str {
    strip_generic_application_text(type_name).trim()
}

pub fn first_generic_argument_leaf_name(type_name: &str) -> Option<String> {
    generic_argument_type_refs(type_name)
        .into_iter()
        .next()
        .map(|ty| {
            let display = display_type_ref(&ty);
            display
                .rsplit('.')
                .next()
                .unwrap_or(display.as_str())
                .trim()
                .to_string()
        })
}

pub fn open_generic_type_name(type_name: &str) -> String {
    let Some(TypeRef {
        kind: TypeRefKind::Named { path, args } }) = parse_type_ref_hint(type_name)
    else {
        return erased_type_name(type_name);
    };
    if args.is_empty() {
        path.display_name()
    } else {
        format!(
            "{}<{}>",
            path.display_name(),
            vec!["_"; args.len()].join(", ")
        )
    }
}

pub fn generic_argument_type_refs(type_name: &str) -> Vec<TypeRef> {
    if let Some(args) = parse_generic_argument_list(type_name) {
        return args;
    }
    match parse_type_ref_hint(type_name) {
        Some(TypeRef {
            kind: TypeRefKind::Named { args, .. } }) => args
            .into_iter()
            .filter_map(|arg| match arg {
                GenericArg::Type(ty) => Some(ty),
                GenericArg::Const(_) => None })
            .collect(),
        _ => Vec::new() }
}

pub fn generic_argument_display_names(type_name: &str) -> Vec<String> {
    generic_argument_type_refs(type_name)
        .into_iter()
        .map(|ty| display_type_ref(&ty))
        .collect()
}

pub fn erase_type(type_ref: &TypeRef) -> TypeRef {
    match &type_ref.kind {
        TypeRefKind::Named { path, .. } => TypeRef {
            kind: TypeRefKind::Named {
                path: path.clone(),
                args: Vec::new() } },
        TypeRefKind::Array { element, rank } => TypeRef {
            kind: TypeRefKind::Array {
                element: Box::new(erase_type(element)),
                rank: *rank } },
        TypeRefKind::Nullable { inner } => TypeRef {
            kind: TypeRefKind::Nullable {
                inner: Box::new(erase_type(inner)) } },
        TypeRefKind::Pointer { inner } => TypeRef {
            kind: TypeRefKind::Pointer {
                inner: Box::new(erase_type(inner)) } },
        TypeRefKind::Reference { inner } => TypeRef {
            kind: TypeRefKind::Reference {
                inner: Box::new(erase_type(inner)) } },
        TypeRefKind::Tuple { elements } => TypeRef {
            kind: TypeRefKind::Tuple {
                elements: elements.iter().map(erase_type).collect() } },
        TypeRefKind::Function { params, result } => TypeRef {
            kind: TypeRefKind::Function {
                params: params.iter().map(erase_type).collect(),
                result: Box::new(erase_type(result)) } },
        TypeRefKind::Union { members } => TypeRef {
            kind: TypeRefKind::Union {
                members: members.iter().map(erase_type).collect() } },
        TypeRefKind::Intersection { members } => TypeRef {
            kind: TypeRefKind::Intersection {
                members: members.iter().map(erase_type).collect() } },
        TypeRefKind::Wildcard { .. }
        | TypeRefKind::GenericParam { .. }
        | TypeRefKind::SelfType
        | TypeRefKind::Infer
        | TypeRefKind::Error => type_ref.clone() }
}

pub fn substitute_type(type_ref: &TypeRef, ctx: &GenericContext) -> TypeRef {
    match &type_ref.kind {
        TypeRefKind::GenericParam { name } => {
            ctx.get(name).cloned().unwrap_or_else(|| type_ref.clone())
        }
        TypeRefKind::Named { path, args } => TypeRef {
            kind: TypeRefKind::Named {
                path: path.clone(),
                args: args.iter().map(|arg| substitute_arg(arg, ctx)).collect() } },
        TypeRefKind::Array { element, rank } => TypeRef {
            kind: TypeRefKind::Array {
                element: Box::new(substitute_type(element, ctx)),
                rank: *rank } },
        TypeRefKind::Tuple { elements } => TypeRef {
            kind: TypeRefKind::Tuple {
                elements: elements.iter().map(|ty| substitute_type(ty, ctx)).collect() } },
        TypeRefKind::Function { params, result } => TypeRef {
            kind: TypeRefKind::Function {
                params: params.iter().map(|ty| substitute_type(ty, ctx)).collect(),
                result: Box::new(substitute_type(result, ctx)) } },
        TypeRefKind::Union { members } => TypeRef {
            kind: TypeRefKind::Union {
                members: members.iter().map(|ty| substitute_type(ty, ctx)).collect() } },
        TypeRefKind::Intersection { members } => TypeRef {
            kind: TypeRefKind::Intersection {
                members: members.iter().map(|ty| substitute_type(ty, ctx)).collect() } },
        TypeRefKind::Nullable { inner } => TypeRef {
            kind: TypeRefKind::Nullable {
                inner: Box::new(substitute_type(inner, ctx)) } },
        TypeRefKind::Pointer { inner } => TypeRef {
            kind: TypeRefKind::Pointer {
                inner: Box::new(substitute_type(inner, ctx)) } },
        TypeRefKind::Reference { inner } => TypeRef {
            kind: TypeRefKind::Reference {
                inner: Box::new(substitute_type(inner, ctx)) } },
        TypeRefKind::Wildcard { bound } => TypeRef {
            kind: TypeRefKind::Wildcard {
                bound: bound.as_ref().map(|bound| substitute_bound(bound, ctx)) } },
        TypeRefKind::SelfType | TypeRefKind::Infer | TypeRefKind::Error => type_ref.clone() }
}

pub fn check_arity(args: &[TypeRef], params: &[GenericParam]) -> Result<(), GenericError> {
    let required = params
        .iter()
        .filter(|param| param.default.is_none())
        .count();
    if args.len() < required || args.len() > params.len() {
        return Err(GenericError::ArityMismatch {
            expected: params.len(),
            actual: args.len() });
    }
    Ok(())
}

pub fn display_type_ref(type_ref: &TypeRef) -> String {
    match &type_ref.kind {
        TypeRefKind::Named { path, args } => {
            let base = path.display_name();
            if args.is_empty() {
                base
            } else {
                format!(
                    "{}<{}>",
                    base,
                    args.iter()
                        .map(display_generic_arg)
                        .collect::<Vec<_>>()
                        .join(", ")
                )
            }
        }
        TypeRefKind::GenericParam { name } => name.clone(),
        TypeRefKind::Array { element, rank } => {
            if *rank <= 1 {
                format!("{}[]", display_type_ref(element))
            } else {
                format!("{}[{}]", display_type_ref(element), ",".repeat(rank - 1))
            }
        }
        TypeRefKind::Tuple { elements } => format!(
            "({})",
            elements
                .iter()
                .map(display_type_ref)
                .collect::<Vec<_>>()
                .join(", ")
        ),
        TypeRefKind::Function { params, result } => format!(
            "fn({}) -> {}",
            params
                .iter()
                .map(display_type_ref)
                .collect::<Vec<_>>()
                .join(", "),
            display_type_ref(result)
        ),
        TypeRefKind::Union { members } => members
            .iter()
            .map(display_type_ref)
            .collect::<Vec<_>>()
            .join(" | "),
        TypeRefKind::Intersection { members } => members
            .iter()
            .map(display_type_ref)
            .collect::<Vec<_>>()
            .join(" & "),
        TypeRefKind::Nullable { inner } => format!("{}?", display_type_ref(inner)),
        TypeRefKind::Pointer { inner } => format!("*{}", display_type_ref(inner)),
        TypeRefKind::Reference { inner } => format!("&{}", display_type_ref(inner)),
        TypeRefKind::Wildcard { bound: None } => "?".to_string(),
        TypeRefKind::Wildcard {
            bound: Some(GenericBound::Extends(ty)) } => format!("? extends {}", display_type_ref(ty)),
        TypeRefKind::Wildcard {
            bound: Some(GenericBound::Super(ty)) } => format!("? super {}", display_type_ref(ty)),
        TypeRefKind::SelfType => "Self".to_string(),
        TypeRefKind::Infer => "_".to_string(),
        TypeRefKind::Error => "<error>".to_string() }
}

pub fn parse_type_ref_hint(text: &str) -> Option<TypeRef> {
    parse_type_ref_inner(text.trim())
}

pub fn parse_generic_params_hint(text: &str) -> Vec<GenericParam> {
    let Some(inner) = generic_params_inner(text) else {
        return Vec::new();
    };
    split_top_level(inner, ',')
        .into_iter()
        .filter_map(parse_generic_param)
        .collect()
}

fn substitute_arg(arg: &GenericArg, ctx: &GenericContext) -> GenericArg {
    match arg {
        GenericArg::Type(ty) => GenericArg::Type(substitute_type(ty, ctx)),
        GenericArg::Const(value) => GenericArg::Const(value.clone()) }
}

fn substitute_bound(bound: &GenericBound, ctx: &GenericContext) -> GenericBound {
    match bound {
        GenericBound::Extends(ty) => GenericBound::Extends(Box::new(substitute_type(ty, ctx))),
        GenericBound::Super(ty) => GenericBound::Super(Box::new(substitute_type(ty, ctx))) }
}

fn display_generic_arg(arg: &GenericArg) -> String {
    match arg {
        GenericArg::Type(ty) => display_type_ref(ty),
        GenericArg::Const(value) => value.clone() }
}

fn generic_params_inner(text: &str) -> Option<&str> {
    let trimmed = text.trim();
    if trimmed.starts_with('<') && trimmed.ends_with('>') {
        return Some(&trimmed[1..trimmed.len() - 1]);
    }
    if trimmed.starts_with('(') && trimmed.ends_with(')') {
        let inner = trimmed[1..trimmed.len() - 1].trim();
        return inner
            .strip_prefix("Of ")
            .or_else(|| inner.strip_prefix("of "))
            .or_else(|| inner.strip_prefix("OF "));
    }
    if trimmed.starts_with('[') && trimmed.ends_with(']') {
        return Some(&trimmed[1..trimmed.len() - 1]);
    }
    trimmed
        .strip_prefix("Of ")
        .or_else(|| trimmed.strip_prefix("of "))
        .or_else(|| trimmed.strip_prefix("OF "))
}

fn parse_generic_param(text: &str) -> Option<GenericParam> {
    let mut trimmed = text.trim();
    if trimmed.is_empty() {
        return None;
    }

    let mut variance = GenericVariance::Invariant;
    if trimmed.len() >= 4 && trimmed[..4].eq_ignore_ascii_case("out ") {
        variance = GenericVariance::Covariant;
        trimmed = trimmed[4..].trim();
    } else if trimmed.len() >= 3 && trimmed[..3].eq_ignore_ascii_case("in ") {
        variance = GenericVariance::Contravariant;
        trimmed = trimmed[3..].trim();
    } else if let Some(rest) = trimmed.strip_prefix("+") {
        variance = GenericVariance::Covariant;
        trimmed = rest.trim();
    } else if let Some(rest) = trimmed.strip_prefix("-") {
        variance = GenericVariance::Contravariant;
        trimmed = rest.trim();
    }

    let (without_default, default) = split_top_level_once(trimmed, '=');
    let default = default.and_then(parse_type_ref_inner);
    let (name_part, constraints_text) = split_generic_param_constraints(without_default);
    let name = name_part
        .split_whitespace()
        .next()
        .map(|name| {
            name.trim_matches(|ch: char| {
                !ch.is_alphanumeric() && ch != '_' && ch != '$' && ch != '@'
            })
        })
        .filter(|name| !name.is_empty())?
        .to_string();

    let constraints = constraints_text
        .map(parse_generic_constraints)
        .unwrap_or_default();
    Some(GenericParam {
        name,
        constraints,
        variance,
        default,
        runtime: GenericRuntimeMode::Erased })
}

fn split_generic_param_constraints(text: &str) -> (&str, Option<&str>) {
    let trimmed = text.trim();
    let lower = trimmed.to_ascii_lowercase();
    if let Some(index) = lower.find(" extends ") {
        return (
            &trimmed[..index],
            Some(&trimmed[index + " extends ".len()..]),
        );
    }
    if let Some(index) = lower.find(" super ") {
        return (&trimmed[..index], Some(&trimmed[index + " super ".len()..]));
    }
    if let Some(index) = lower.find(" as ") {
        return (&trimmed[..index], Some(&trimmed[index + " as ".len()..]));
    }
    let (left, right) = split_top_level_once(trimmed, ':');
    if let Some(right) = right {
        return (left, Some(right));
    }
    if let Some((index, _)) = trimmed.char_indices().find(|(_, ch)| ch.is_whitespace()) {
        let left = trimmed[..index].trim();
        let right = trimmed[index..].trim();
        if !left.is_empty() && !right.is_empty() {
            return (left, Some(right));
        }
    }
    (trimmed, None)
}

fn parse_generic_constraints(text: &str) -> Vec<GenericConstraint> {
    let mut trimmed = text.trim();
    if trimmed.starts_with('{') && trimmed.ends_with('}') {
        trimmed = trimmed[1..trimmed.len() - 1].trim();
    }
    split_constraint_list(trimmed)
        .into_iter()
        .filter_map(parse_generic_constraint)
        .collect()
}

fn parse_generic_constraint(text: &str) -> Option<GenericConstraint> {
    let trimmed = text.trim().trim_start_matches('~').trim();
    if trimmed.is_empty() {
        return None;
    }
    let lower = trimmed.to_ascii_lowercase();
    match lower.as_str() {
        "any" | "object" => Some(GenericConstraint::Any),
        "class" => Some(GenericConstraint::Class),
        "struct" | "record" => Some(GenericConstraint::Struct),
        "interface" => Some(GenericConstraint::Interface),
        "enum" => Some(GenericConstraint::Enum),
        "delegate" => Some(GenericConstraint::Delegate),
        "new" | "new()" | "constructor" | "create" => {
            Some(GenericConstraint::Constructor { argc: None })
        }
        "notnull" | "nonnull" => Some(GenericConstraint::NonNull),
        "nullable" => Some(GenericConstraint::Nullable),
        "unmanaged" => Some(GenericConstraint::Unmanaged),
        "comparable" => Some(GenericConstraint::Comparable),
        "number" | "numeric" => Some(GenericConstraint::Numeric),
        "integer" | "int" => Some(GenericConstraint::Integer),
        "float" | "floating" | "double" => Some(GenericConstraint::Floating),
        _ => parse_type_ref_inner(trimmed).map(GenericConstraint::Extends) }
}

fn split_constraint_list(text: &str) -> Vec<&str> {
    split_top_level(text, ',')
        .into_iter()
        .flat_map(|part| split_top_level(part, '&'))
        .flat_map(|part| split_top_level(part, '|'))
        .collect()
}

fn parse_type_ref_inner(text: &str) -> Option<TypeRef> {
    let trimmed = text.trim();
    if trimmed.is_empty() {
        return None;
    }

    if let Some(array_type) = parse_pascal_array_type(trimmed) {
        return Some(array_type);
    }

    if trimmed == "?" {
        return Some(TypeRef {
            kind: TypeRefKind::Wildcard { bound: None } });
    }
    if let Some(rest) = trimmed.strip_prefix("? extends ") {
        return Some(TypeRef {
            kind: TypeRefKind::Wildcard {
                bound: Some(GenericBound::Extends(Box::new(parse_type_ref_inner(rest)?))) } });
    }
    if let Some(rest) = trimmed.strip_prefix("? super ") {
        return Some(TypeRef {
            kind: TypeRefKind::Wildcard {
                bound: Some(GenericBound::Super(Box::new(parse_type_ref_inner(rest)?))) } });
    }
    if let Some(inner) = trimmed.strip_suffix('?') {
        return Some(TypeRef {
            kind: TypeRefKind::Nullable {
                inner: Box::new(parse_type_ref_inner(inner)?) } });
    }
    if let Some(inner) = trimmed.strip_suffix("[]") {
        return Some(TypeRef {
            kind: TypeRefKind::Array {
                element: Box::new(parse_type_ref_inner(inner)?),
                rank: 1 } });
    }

    if let Some((base, args)) = split_generic_application(trimmed) {
        let parsed_args = parse_generic_args(args)?;
        return Some(TypeRef {
            kind: TypeRefKind::Named {
                path: TypePath::from_dotted(base.trim()),
                args: parsed_args } });
    }

    if let Some((base, args)) = split_vb_generic_application(trimmed) {
        let parsed_args = parse_generic_args(args)?;
        return Some(TypeRef {
            kind: TypeRefKind::Named {
                path: TypePath::from_dotted(base.trim()),
                args: parsed_args } });
    }

    if let Some((base, args)) = split_square_generic_application(trimmed) {
        let parsed_args = parse_generic_args(args)?;
        return Some(TypeRef {
            kind: TypeRefKind::Named {
                path: TypePath::from_dotted(base.trim()),
                args: parsed_args } });
    }

    Some(TypeRef {
        kind: TypeRefKind::Named {
            path: TypePath::from_dotted(trimmed),
            args: Vec::new() } })
}

fn parse_generic_argument_list(text: &str) -> Option<Vec<TypeRef>> {
    let trimmed = text.trim();
    let args = if trimmed.starts_with('<') && trimmed.ends_with('>') {
        &trimmed[1..trimmed.len() - 1]
    } else if trimmed.starts_with('[') && trimmed.ends_with(']') {
        &trimmed[1..trimmed.len() - 1]
    } else if trimmed.starts_with('(') && trimmed.ends_with(')') {
        let inner = trimmed[1..trimmed.len() - 1].trim();
        inner
            .strip_prefix("Of ")
            .or_else(|| inner.strip_prefix("of "))
            .or_else(|| inner.strip_prefix("OF "))?
    } else {
        return None;
    };
    split_top_level(args, ',')
        .into_iter()
        .map(parse_type_ref_inner)
        .collect()
}

fn parse_generic_args(args: &str) -> Option<Vec<GenericArg>> {
    if args.trim().is_empty() {
        return Some(Vec::new());
    }
    split_top_level(args, ',')
        .into_iter()
        .map(|arg| parse_type_ref_inner(arg).map(GenericArg::Type))
        .collect()
}

fn parse_pascal_array_type(text: &str) -> Option<TypeRef> {
    let lower = text.to_ascii_lowercase();
    if !lower.starts_with("array") {
        return None;
    }
    let of_index = lower.rfind(" of ")?;
    let element = parse_type_ref_inner(&text[of_index + 4..])?;
    let prefix = text[..of_index].trim();
    let rank = if let (Some(start), Some(end)) = (prefix.find('['), prefix.rfind(']')) {
        prefix[start + 1..end]
            .split(',')
            .filter(|dim| !dim.trim().is_empty())
            .count()
            .max(1)
    } else {
        1
    };
    Some(TypeRef {
        kind: TypeRefKind::Array {
            element: Box::new(element),
            rank } })
}

fn split_generic_application(text: &str) -> Option<(&str, &str)> {
    let start = text.find('<')?;
    if !text.ends_with('>') {
        return None;
    }
    if matching_angle_end(text, start)? != text.len() - 1 {
        return None;
    }
    Some((&text[..start], &text[start + 1..text.len() - 1]))
}

fn split_vb_generic_application(text: &str) -> Option<(&str, &str)> {
    let lower = text.to_ascii_lowercase();
    let start = lower.find("(of ")?;
    if !text.ends_with(')') {
        return None;
    }
    Some((text[..start].trim(), text[start + 4..text.len() - 1].trim()))
}

fn split_square_generic_application(text: &str) -> Option<(&str, &str)> {
    let start = text.find('[')?;
    if start == 0 || !text.ends_with(']') || text.ends_with("[]") {
        return None;
    }
    if matching_bracket_end(text, start, '[', ']')? != text.len() - 1 {
        return None;
    }
    Some((&text[..start], &text[start + 1..text.len() - 1]))
}

fn strip_generic_application_text(text: &str) -> &str {
    let trimmed = text.trim();
    let angle = trimmed.find('<');
    let vb = trimmed.to_ascii_lowercase().find("(of ");
    let square = split_square_generic_application(trimmed).map(|(base, _)| base.len());
    match (angle, vb, square) {
        (Some(a), Some(b), Some(c)) => &trimmed[..a.min(b).min(c)],
        (Some(a), Some(b), None) => &trimmed[..a.min(b)],
        (Some(a), None, Some(c)) => &trimmed[..a.min(c)],
        (None, Some(b), Some(c)) => &trimmed[..b.min(c)],
        (Some(a), None, None) => &trimmed[..a],
        (None, Some(b), None) => &trimmed[..b],
        (None, None, Some(c)) => &trimmed[..c],
        (None, None, None) => trimmed }
}

fn matching_angle_end(text: &str, start: usize) -> Option<usize> {
    matching_bracket_end(text, start, '<', '>')
}

fn matching_bracket_end(text: &str, start: usize, open: char, close: char) -> Option<usize> {
    let mut depth = 0usize;
    for (index, ch) in text.char_indices().skip_while(|(index, _)| *index < start) {
        if ch == open {
            depth += 1;
        } else if ch == close {
            depth = depth.checked_sub(1)?;
            if depth == 0 {
                return Some(index);
            }
        } else {
            match ch {
                '<' | '(' | '[' | '{' if open != ch => depth += 1,
                '>' | ')' | ']' | '}' if close != ch => {
                    depth = depth.checked_sub(1)?;
                }
                _ => {}
            }
        }
    }
    None
}

fn split_top_level(text: &str, delimiter: char) -> Vec<&str> {
    let mut parts = Vec::new();
    let mut depth = 0usize;
    let mut start = 0usize;
    for (index, ch) in text.char_indices() {
        match ch {
            '<' | '(' | '[' | '{' => depth += 1,
            '>' | ')' | ']' | '}' => depth = depth.saturating_sub(1),
            ch if ch == delimiter && depth == 0 => {
                parts.push(text[start..index].trim());
                start = index + ch.len_utf8();
            }
            _ => {}
        }
    }
    parts.push(text[start..].trim());
    parts
}

fn split_top_level_once(text: &str, delimiter: char) -> (&str, Option<&str>) {
    let mut depth = 0usize;
    for (index, ch) in text.char_indices() {
        match ch {
            '<' | '(' | '[' | '{' => depth += 1,
            '>' | ')' | ']' | '}' => depth = depth.saturating_sub(1),
            ch if ch == delimiter && depth == 0 => {
                return (&text[..index], Some(&text[index + ch.len_utf8()..]));
            }
            _ => {}
        }
    }
    (text, None)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parses_nested_generic_type_hints() {
        let ty = parse_type_ref_hint("Map<String, List<Int>>").unwrap();

        assert_eq!(display_type_ref(&ty), "Map<String, List<Int>>");
        assert_eq!(canonical_type_ref(&ty).name, "Map<String, List<Int>>");
    }

    #[test]
    fn erases_named_generic_args_recursively() {
        let ty = parse_type_ref_hint("Map<String, List<Int>>[]").unwrap();
        let erased = erase_type(&ty);

        assert_eq!(display_type_ref(&erased), "Map[]");
    }

    #[test]
    fn substitutes_generic_params_inside_nested_args() {
        let ty = TypeRef {
            kind: TypeRefKind::Named {
                path: TypePath::from_dotted("Box"),
                args: vec![GenericArg::Type(TypeRef::generic_param("T"))] } };
        let ctx = GenericContext::new().with_substitution("T", TypeRef::named("String"));

        assert_eq!(display_type_ref(&substitute_type(&ty, &ctx)), "Box<String>");
    }

    #[test]
    fn checks_required_and_defaulted_arity() {
        let params = vec![
            GenericParam::new("T"),
            GenericParam {
                default: Some(TypeRef::named("String")),
                ..GenericParam::new("U")
            },
        ];

        assert!(check_arity(&[TypeRef::named("Int")], &params).is_ok());
        assert!(check_arity(&[], &params).is_err());
        assert!(
            check_arity(
                &[
                    TypeRef::named("Int"),
                    TypeRef::named("String"),
                    TypeRef::named("Bool")
                ],
                &params
            )
            .is_err()
        );
    }

    #[test]
    fn binds_defaulted_generic_signature_args() {
        let signature = GenericSignature::new(vec![
            GenericParam::new("T"),
            GenericParam {
                default: Some(TypeRef::named("String")),
                ..GenericParam::new("U")
            },
        ]);
        let ctx = signature.bind_args(&[TypeRef::named("Integer")]).unwrap();

        assert_eq!(
            display_type_ref(ctx.get("T").expect("T substitution")),
            "Integer"
        );
        assert_eq!(
            display_type_ref(ctx.get("U").expect("U substitution")),
            "String"
        );
    }

    #[test]
    fn parses_wildcard_bounds() {
        let ty = parse_type_ref_hint("? extends Number").unwrap();

        assert_eq!(display_type_ref(&ty), "? extends Number");
    }

    #[test]
    fn parses_vb_generic_type_hints() {
        let ty = parse_type_ref_hint("List(Of Integer)").unwrap();

        assert_eq!(display_type_ref(&ty), "List<Integer>");
        assert_eq!(
            generic_argument_display_names("Dictionary(Of String, Integer)"),
            vec!["String", "Integer"]
        );
        assert_eq!(
            generic_argument_display_names("(Of String, Integer)"),
            vec!["String", "Integer"]
        );
    }

    #[test]
    fn parses_pascal_array_type_hints() {
        let ty = parse_type_ref_hint("array[1..3, 1..4] of TList<Integer>").unwrap();

        assert_eq!(display_type_ref(&ty), "TList<Integer>[,]");
    }

    #[test]
    fn parses_generic_param_declarations() {
        let java_params =
            parse_generic_params_hint("<T extends Number & Comparable<T>, U = String>");
        assert_eq!(java_params.len(), 2);
        assert_eq!(java_params[0].name, "T");
        assert_eq!(java_params[0].constraints.len(), 2);
        assert_eq!(java_params[1].name, "U");
        assert_eq!(
            java_params[1]
                .default
                .as_ref()
                .map(display_type_ref)
                .as_deref(),
            Some("String")
        );

        let vb_params = parse_generic_params_hint("(Of Out T As {Class, New}, In U As IFoo)");
        assert_eq!(vb_params.len(), 2);
        assert_eq!(vb_params[0].variance, GenericVariance::Covariant);
        assert_eq!(vb_params[0].constraints.len(), 2);
        assert_eq!(vb_params[1].variance, GenericVariance::Contravariant);
        assert_eq!(vb_params[1].constraints.len(), 1);

        let go_params = parse_generic_params_hint("[T any, U ~int | string]");
        assert_eq!(go_params.len(), 2);
        assert_eq!(go_params[0].name, "T");
        assert_eq!(go_params[0].constraints, vec![GenericConstraint::Any]);
        assert_eq!(go_params[1].name, "U");
        assert_eq!(go_params[1].constraints.len(), 2);
        assert_eq!(
            runtime_type_arg_param_names(&go_params),
            vec!["__generic_typearg_T", "__generic_typearg_U"]
        );
    }

    #[test]
    fn derives_erased_and_open_generic_names() {
        assert_eq!(erased_type_name("Map<String, List<Int>>"), "Map");
        assert_eq!(
            open_generic_type_name("Map<String, List<Int>>"),
            "Map<_, _>"
        );
        assert_eq!(erased_type_name("List(Of Integer)"), "List");
        assert_eq!(
            generic_base_name("Dictionary(Of String, Integer)"),
            "Dictionary"
        );
        assert_eq!(
            first_generic_argument_leaf_name("java.util.List<java.lang.String>").as_deref(),
            Some("String")
        );
        assert_eq!(
            generic_argument_display_names("<String, List<Int>>"),
            vec!["String", "List<Int>"]
        );
        assert_eq!(erased_type_name("Pair[int, string]"), "Pair");
        assert_eq!(
            generic_argument_display_names("Pair[int, []string]"),
            vec!["int", "[]string"]
        );
        assert_eq!(
            generic_argument_display_names("[int, []string]"),
            vec!["int", "[]string"]
        );
        assert!(generic_argument_display_names("<>").is_empty());
        assert_eq!(
            display_type_ref(&parse_type_ref_hint("ArrayList<>").unwrap()),
            "ArrayList"
        );
    }
}
