//! Primary component source -> typed binary AST. No component WAT encoder.
use super::component_source_scope::{Namespace, SourceInterface, SourceScope, SourceShape};
use super::{Pair, Rule, component_name, find_rule, parse_unsigned_const, unquote_to_bytes};
use std::sync::Arc;
#[path = "component_value_mixed.rs"]
mod mixed;
use vybe_ast::canon::{
    ComponentBinary, ComponentBinaryDef as Def, ComponentExternalName as Name,
    ComponentInterfaceDecl as Decl, ComponentInterfaceRef as Ref, ComponentItemKind as Kind,
    ComponentItemRef as Item, ComponentNameAttribute as Attribute,
    ComponentNamedItemRef as NamedItem, ComponentOuterAlias, ComponentOuterAliasKind,
    ValSpec as Val,
};

pub(super) fn contains(source: Pair<'_, Rule>) -> bool {
    matches!(
        source.as_rule(),
        Rule::comp_value | Rule::comp_start | Rule::valuebound
            | Rule::canon_spawn_ref | Rule::canon_spawn_indirect | Rule::canon_avail_par
    ) || source
        .clone()
        .into_inner()
        .any(|p| p.as_rule() == Rule::comp_sort_kw && p.as_str() == "value" || contains(p))
}

pub(super) fn concrete(source: Pair<'_, Rule>) -> bool {
    if matches!(source.as_rule(), Rule::instancetype | Rule::componenttype) {
        return false;
    }
    matches!(
        source.as_rule(),
        Rule::comp_value | Rule::comp_start | Rule::valuebound
            | Rule::canon_spawn_ref | Rule::canon_spawn_indirect | Rule::canon_avail_par
    ) || source.as_rule() == Rule::comp_sort_kw && source.as_str() == "value"
        || source.into_inner().any(concrete)
}

fn child<'a>(source: &Pair<'a, Rule>, rule: Rule) -> Result<Pair<'a, Rule>, String> {
    source
        .clone()
        .into_inner()
        .find(|p| p.as_rule() == rule)
        .ok_or_else(|| format!("component value text: missing {rule:?}"))
}

fn binder(source: &Pair<'_, Rule>) -> Option<String> {
    source
        .clone()
        .into_inner()
        .find(|p| p.as_rule() == Rule::id)
        .map(|p| p.as_str()[1..].to_owned())
}

fn head<'a>(source: &Pair<'a, Rule>) -> &'a str {
    // Keywords are fixed by the grammar; skip trivia before the keyword too.
    let mut s = source
        .as_str()
        .trim_start()
        .strip_prefix('(')
        .unwrap_or(source.as_str());
    loop {
        s = s.trim_start();
        if let Some(rest) = s.strip_prefix(";;") {
            s = rest.split_once('\n').map_or("", |(_, tail)| tail);
        } else if s.starts_with("(;") {
            let mut depth = 1;
            let mut i = 2;
            while depth != 0 && i + 1 < s.len() {
                match &s.as_bytes()[i..i + 2] {
                    b"(;" => {
                        depth += 1;
                        i += 2;
                    }
                    b";)" => {
                        depth -= 1;
                        i += 2;
                    }
                    _ => i += 1,
                }
            }
            s = &s[i..];
        } else {
            break;
        }
    }
    s.split(|c: char| c.is_whitespace() || c == ')')
        .next()
        .unwrap_or("")
}

fn slot(kind: Kind) -> usize {
    match kind {
        Kind::Module => 0,
        Kind::Func => 1,
        Kind::Value => 2,
        Kind::Type => 3,
        Kind::Component => 4,
        Kind::Instance => 5,
    }
}

fn kind(text: &str) -> Result<Kind, String> {
    match text {
        "core module" | "module" => Ok(Kind::Module),
        "func" => Ok(Kind::Func),
        "value" => Ok(Kind::Value),
        "type" => Ok(Kind::Type),
        "component" => Ok(Kind::Component),
        "instance" => Ok(Kind::Instance),
        _ => Err(format!("unsupported component item sort {text}")),
    }
}

#[derive(Default)]
struct Producer {
    scope: SourceScope,
    outers: Vec<SourceScope>,
    binary: ComponentBinary,
    pending_interface: Option<SourceInterface>,
}

impl Producer {
    fn type_index(&mut self, source: &Pair<'_, Rule>) -> Result<u32, String> {
        if !source.as_str().starts_with('$')
            || self.scope.names[slot(Kind::Type)].contains_key(&source.as_str()[1..])
        {
            return self.scope.index(Kind::Type, source);
        }
        for (position, outer) in self.outers.iter().enumerate().rev() {
            if let Some(&index) = outer.names[slot(Kind::Type)].get(&source.as_str()[1..]) {
                let item = outer.item(Namespace::Type, index)?;
                let count = u32::try_from(self.outers.len() - position)
                    .map_err(|_| "outer depth overflow")?;
                self.binary
                    .definitions
                    .push(Def::AliasOuter(ComponentOuterAlias {
                        kind: ComponentOuterAliasKind::Type,
                        count,
                        index,
                    }));
                let local = self.scope.alias(
                    Namespace::Type,
                    Some(source.as_str()[1..].to_owned()),
                    &item,
                )?;
                return Ok(local);
            }
        }
        Err(format!("unknown component type {}", source.as_str()))
    }

    fn expand(&self, ty: &Val) -> Result<Val, String> {
        // Every stored shape was expanded before insertion, so one lookup suffices.
        if let Val::Ref(index) = ty {
            let source = self
                .scope
                .projection(Namespace::Type, *index)
                .ok_or("literal type projection is unavailable")?;
            return match &source.shape {
                SourceShape::Value(ty) => Ok(ty.clone()),
                _ => Err("literal type is not a value type".into()),
            };
        }
        Ok(ty.clone())
    }

    fn add_type(&mut self, definition: Def, name: Option<String>) -> Result<u32, String> {
        let shape = match &definition {
            Def::ValueType(ty) => SourceShape::Value(self.shape(ty)?),
            Def::InterfaceType { .. } => SourceShape::Interface(
                self.pending_interface
                    .take()
                    .ok_or("interface source projection is unavailable")?,
            ),
            _ => SourceShape::Opaque,
        };
        self.binary.definitions.push(definition);
        let index = self.scope.add(Kind::Type, name)?;
        self.scope.define_shape(Namespace::Type, index, shape)?;
        Ok(index)
    }

    fn shape(&self, ty: &Val) -> Result<Val, String> {
        let boxed = |v: &Val| self.shape(v).map(Box::new);
        let optional = |v: &Option<Box<Val>>| v.as_deref().map(boxed).transpose();
        Ok(match ty {
            Val::Ref(_) => self.expand(ty)?,
            Val::List(v) => Val::List(boxed(v)?),
            Val::ListFixed(v, n) => Val::ListFixed(boxed(v)?, *n),
            Val::Option(v) => Val::Option(boxed(v)?),
            Val::Result(a, b) => Val::Result(optional(a)?, optional(b)?),
            Val::Map(a, b) => Val::Map(boxed(a)?, boxed(b)?),
            Val::Future(v) => Val::Future(optional(v)?),
            Val::Stream(v) => Val::Stream(optional(v)?),
            Val::Tuple(v) => Val::Tuple(v.iter().map(|v| self.shape(v)).collect::<Result<_, _>>()?),
            Val::Record(v) => Val::Record(
                v.iter()
                    .map(|(n, v)| Ok((n.clone(), self.shape(v)?)))
                    .collect::<Result<_, String>>()?,
            ),
            Val::Variant(v) => Val::Variant(
                v.iter()
                    .map(|(n, v)| Ok((n.clone(), v.as_ref().map(|v| self.shape(v)).transpose()?)))
                    .collect::<Result<_, String>>()?,
            ),
            _ => ty.clone(),
        })
    }

    fn valtype(&mut self, source: Pair<'_, Rule>) -> Result<Val, String> {
        let inner = source.into_inner().next().ok_or("empty value type")?;
        if inner.as_rule() == Rule::index {
            return self.type_index(&inner).map(Val::Ref);
        }
        let ty = self.defvaltype(inner)?;
        if matches!(ty, Val::Prim(_)) {
            Ok(ty)
        } else {
            self.add_type(Def::ValueType(ty), None).map(Val::Ref)
        }
    }

    fn defvaltype(&mut self, source: Pair<'_, Rule>) -> Result<Val, String> {
        let tag = head(&source).to_owned();
        if !source.as_str().trim_start().starts_with('(') {
            return Ok(Val::Prim(source.as_str().into()));
        }
        let parts: Vec<_> = source.into_inner().collect();
        let mut vals = Vec::new();
        for part in &parts {
            if part.as_rule() == Rule::valtype {
                vals.push(self.valtype(part.clone())?);
            }
        }
        let first = || {
            vals.first()
                .cloned()
                .map(Box::new)
                .ok_or("missing element type".to_string())
        };
        Ok(match tag.as_str() {
            "list" => {
                if let Some(n) = parts.iter().find(|p| p.as_rule() == Rule::integer) {
                    Val::ListFixed(
                        first()?,
                        u32::try_from(parse_unsigned_const(n.as_str())?)
                            .map_err(|_| "list length overflow")?,
                    )
                } else {
                    Val::List(first()?)
                }
            }
            "tuple" => Val::Tuple(vals),
            "option" => Val::Option(first()?),
            "stream" => Val::Stream(vals.first().cloned().map(Box::new)),
            "future" => Val::Future(vals.first().cloned().map(Box::new)),
            "result" => {
                let err = parts
                    .iter()
                    .find(|p| p.as_rule() == Rule::result_err)
                    .map(|p| self.valtype(child(p, Rule::valtype)?).map(Box::new))
                    .transpose()?;
                Val::Result(vals.first().cloned().map(Box::new), err)
            }
            "map" => {
                let key = parts
                    .iter()
                    .find(|p| p.as_rule() == Rule::keytype)
                    .ok_or("missing map key")?;
                Val::Map(Box::new(Val::Prim(key.as_str().into())), first()?)
            }
            "own" | "borrow" => {
                let index = self.type_index(
                    parts
                        .iter()
                        .find(|p| p.as_rule() == Rule::index)
                        .ok_or("missing resource index")?,
                )?;
                if tag == "own" {
                    Val::Own(index)
                } else {
                    Val::Borrow(index)
                }
            }
            "flags" | "enum" => {
                let names = parts
                    .iter()
                    .filter(|p| p.as_rule() == Rule::labellit)
                    .map(|p| component_name(p.as_str(), "literal label"))
                    .collect::<Result<_, _>>()?;
                if tag == "flags" {
                    Val::Flags(names)
                } else {
                    Val::Enum(names)
                }
            }
            "record" => Val::Record(
                parts
                    .iter()
                    .map(|p| {
                        Ok((
                            component_name(child(p, Rule::labellit)?.as_str(), "field")?,
                            self.valtype(child(p, Rule::valtype)?)?,
                        ))
                    })
                    .collect::<Result<_, String>>()?,
            ),
            "variant" => Val::Variant(
                parts
                    .iter()
                    .map(|p| {
                        Ok((
                            component_name(child(p, Rule::labellit)?.as_str(), "case")?,
                            p.clone()
                                .into_inner()
                                .find(|p| p.as_rule() == Rule::valtype)
                                .map(|p| self.valtype(p))
                                .transpose()?,
                        ))
                    })
                    .collect::<Result<_, String>>()?,
            ),
            _ => return Err(format!("unsupported value type {tag}")),
        })
    }

    fn typedef(&mut self, source: Pair<'_, Rule>) -> Result<Def, String> {
        Ok(match source.as_rule() {
            Rule::defvaltype => Def::ValueType(self.defvaltype(source)?),
            Rule::functype_c => {
                let mut params = Vec::new();
                let mut result = None;
                let mut is_async = false;
                for p in source.into_inner() {
                    match p.as_rule() {
                        Rule::func_param_c => params.push((
                            component_name(child(&p, Rule::labellit)?.as_str(), "parameter")?,
                            self.valtype(child(&p, Rule::valtype)?)?,
                        )),
                        Rule::func_result_c => {
                            result = Some(self.valtype(child(&p, Rule::valtype)?)?)
                        }
                        Rule::async_kw => is_async = true,
                        Rule::id => (),
                        _ => return Err("unexpected function type field".into()),
                    }
                }
                Def::FuncType {
                    params,
                    result,
                    is_async,
                }
            }
            Rule::resourcetype => {
                let rep = match child(&source, Rule::resource_rep)?.as_str() {
                    "i32" => vybe_ast::canon::ResourceRep::I32,
                    "i64" => vybe_ast::canon::ResourceRep::I64,
                    _ => return Err("invalid resource representation".into()),
                };
                let dtor = find_rule(&source, Rule::resource_dtor)
                    .map(|p| {
                        self.core_reference(child(&p, Rule::core_funcidx_ref)?, Namespace::CoreFunc)
                    })
                    .transpose()?;
                Def::ResourceType { rep, dtor }
            }
            Rule::instancetype | Rule::componenttype => {
                let is_component = source.as_rule() == Rule::componenttype;
                let mut outers = self.outers.clone();
                outers.push(self.scope.clone());
                let mut nested = Producer {
                    scope: SourceScope::new(&source),
                    outers,
                    ..Producer::default()
                };
                let mut declarations = Vec::new();
                for wrapped in source.into_inner() {
                    if wrapped.as_rule() == Rule::id {
                        continue;
                    }
                    let mut p = wrapped
                        .into_inner()
                        .next()
                        .ok_or("empty interface declaration")?;
                    if p.as_rule() == Rule::instancedecl {
                        p = p.into_inner().next().ok_or("empty instance declaration")?;
                    }
                    let span = p.as_span();
                    nested.scope.site = (span.start(), span.end());
                    let start = nested.binary.definitions.len();
                    if matches!(p.as_rule(), Rule::importdecl | Rule::exportdecl) {
                        let name = external_name(&p)?;
                        let ext = child(&p, Rule::externtype)?;
                        let (ty, bind) = nested.external(ext)?;
                        let index = nested.introduce(&ty, bind)?;
                        if p.as_rule() == Rule::exportdecl {
                            nested.scope.publish(
                                name.name.clone(),
                                external_namespace(&ty),
                                index,
                            )?;
                        }
                        declarations.extend(
                            nested
                                .binary
                                .definitions
                                .drain(start..)
                                .map(interface_decl)
                                .collect::<Result<Vec<_>, _>>()?,
                        );
                        declarations.push(if p.as_rule() == Rule::importdecl {
                            Decl::NamedImport { name, ty }
                        } else {
                            Decl::NamedExport { name, ty }
                        });
                    } else {
                        nested.definition(p)?;
                        declarations.extend(
                            nested
                                .binary
                                .definitions
                                .drain(start..)
                                .map(interface_decl)
                                .collect::<Result<Vec<_>, _>>()?,
                        );
                    }
                }
                nested.scope.check()?;
                self.pending_interface = Some(nested.scope.interface(is_component));
                Def::InterfaceType {
                    is_component,
                    declarations,
                }
            }
            _ => {
                return Err(format!(
                    "value text producer: unsupported type {:?}",
                    source.as_rule()
                ));
            }
        })
    }

    fn external(&mut self, source: Pair<'_, Rule>) -> Result<(Ref, Option<String>), String> {
        if let Some(result) = self.module_external(&source)? {
            return Ok(result);
        }
        let bind = binder(&source);
        if let Some(bound) = source
            .clone()
            .into_inner()
            .find(|p| p.as_rule() == Rule::valuebound)
        {
            let ty = if let Some(v) = bound
                .clone()
                .into_inner()
                .find(|p| p.as_rule() == Rule::valtype)
            {
                Ref::Value(self.valtype(v)?)
            } else {
                Ref::ValueEq(
                    self.scope
                        .index(Kind::Value, &child(&bound, Rule::index)?)?,
                )
            };
            return Ok((ty, bind));
        }
        if let Some(bound) = source
            .clone()
            .into_inner()
            .find(|p| p.as_rule() == Rule::typebound)
        {
            let ty = if let Some(i) = bound.into_inner().find(|p| p.as_rule() == Rule::index) {
                Ref::TypeEq(self.type_index(&i)?)
            } else {
                Ref::TypeSubResource
            };
            return Ok((ty, bind));
        }
        for p in source.clone().into_inner() {
            if matches!(
                p.as_rule(),
                Rule::functype_c | Rule::instancetype | Rule::componenttype
            ) {
                let bind = binder(&p);
                let k = match p.as_rule() {
                    Rule::functype_c => Kind::Func,
                    Rule::instancetype => Kind::Instance,
                    _ => Kind::Component,
                };
                let ty = self.typedef(p)?;
                let index = self.add_type(ty, None)?;
                return Ok((
                    match k {
                        Kind::Func => Ref::Func(index),
                        Kind::Instance => Ref::Instance(index),
                        _ => Ref::Component(index),
                    },
                    bind,
                ));
            }
        }
        let sort = source
            .clone()
            .into_inner()
            .find(|p| p.as_rule() == Rule::comp_sort_kw)
            .map(|p| p.as_str().to_owned())
            .ok_or("unsupported external type in value component")?;
        let index = self.type_index(&child(&source, Rule::index)?)?;
        Ok((
            match sort.as_str() {
                "func" => Ref::Func(index),
                "instance" => Ref::Instance(index),
                "component" => Ref::Component(index),
                _ => return Err("unsupported external type sort".into()),
            },
            bind,
        ))
    }

    fn introduce(&mut self, ty: &Ref, name: Option<String>) -> Result<u32, String> {
        let k = match ty {
            Ref::Module(_) => Kind::Module,
            Ref::Func(_) => Kind::Func,
            Ref::Value(_) | Ref::ValueEq(_) => Kind::Value,
            Ref::TypeEq(_) | Ref::TypeSubResource => Kind::Type,
            Ref::Component(_) => Kind::Component,
            Ref::Instance(_) => Kind::Instance,
        };
        let projection = match ty {
            Ref::TypeEq(i) | Ref::Instance(i) | Ref::Component(i) | Ref::Func(i) => {
                self.scope.projection(Namespace::Type, *i)
            }
            _ => None,
        };
        let origin = if let Ref::TypeEq(i) = ty {
            Some(self.scope.origin(Namespace::Type, *i)?)
        } else {
            None
        };
        let index = self.scope.define(k.into(), name, origin)?;
        self.scope.project(k.into(), index, projection)?;
        if matches!(ty, Ref::TypeSubResource) {
            self.scope
                .define_shape(Namespace::Type, index, SourceShape::Opaque)?;
        }
        Ok(index)
    }

    fn item(&mut self, source: Pair<'_, Rule>) -> Result<(Kind, u32), String> {
        let parts: Vec<_> = source.into_inner().collect();
        let sort = parts.first().ok_or("missing item sort")?;
        let k = kind(sort.as_str())?;
        let reference = parts.get(1).ok_or("missing item index")?;
        let paths: Vec<_> = parts.iter().skip(2).collect();
        if paths.is_empty() {
            return Ok((
                k,
                if k == Kind::Type {
                    self.type_index(reference)?
                } else {
                    self.scope.index(k, reference)?
                },
            ));
        }
        let mut instance = self.scope.index(Kind::Instance, reference)?;
        for (n, path) in paths.iter().enumerate() {
            let target = if n + 1 == paths.len() {
                k
            } else {
                Kind::Instance
            };
            let index = self.alias_export(
                instance,
                component_name(path.as_str(), "export path")?,
                target,
                None,
            )?;
            if n + 1 == paths.len() {
                return Ok((k, index));
            }
            instance = index;
        }
        Err("empty export path".into())
    }

    fn alias_export(
        &mut self,
        instance: u32,
        name: String,
        kind: Kind,
        bind: Option<String>,
    ) -> Result<u32, String> {
        let item = self.scope.export(Namespace::Instance, instance, &name)?;
        let index = self.scope.alias(kind.into(), bind, &item)?;
        self.binary.definitions.push(Def::AliasInstanceExport {
            instance,
            name,
            kind,
        });
        Ok(index)
    }

    fn define_instance(
        &mut self,
        name: Option<String>,
        interface: SourceInterface,
    ) -> Result<u32, String> {
        let index = self.scope.add(Kind::Instance, name)?;
        self.scope.define_shape(
            Namespace::Instance,
            index,
            SourceShape::Interface(SourceInterface {
                is_component: false,
                exports: interface.exports,
            }),
        )?;
        Ok(index)
    }

    fn exported_items(&self, exports: &[NamedItem]) -> Result<SourceInterface, String> {
        let mut interface = SourceInterface::default();
        for item in exports {
            interface.exports.insert(
                item.name.name.clone(),
                self.scope.item(item.kind.into(), item.index)?,
            );
        }
        Ok(interface)
    }

    fn definition(&mut self, source: Pair<'_, Rule>) -> Result<(), String> {
        let span = source.as_span();
        self.scope.site = (span.start(), span.end());
        if self.mixed_definition(source.clone())? {
            return Ok(());
        }
        match source.as_rule() {
            Rule::comp_type => {
                let ty = child(&source, Rule::deftype)?
                    .into_inner()
                    .next()
                    .ok_or("missing type body")?;
                let definition = self.typedef(ty)?;
                self.add_type(definition, binder(&source))?;
            }
            Rule::comp_import => {
                let name = external_name(&source)?;
                let (ty, bind) = self.external(child(&source, Rule::externtype)?)?;
                self.introduce(&ty, bind)?;
                self.binary.definitions.push(Def::NamedImport { name, ty });
            }
            Rule::comp_export => {
                let name = external_name(&source)?;
                let (kind, index) = self.item(child(&source, Rule::externidx)?)?;
                let ty = source
                    .clone()
                    .into_inner()
                    .find(|p| p.as_rule() == Rule::externtype)
                    .map(|p| self.external(p).map(|(ty, _)| ty))
                    .transpose()?;
                let mut item = self.scope.item(kind.into(), index)?;
                if let Some(Ref::Instance(view) | Ref::Component(view)) = &ty {
                    item.projection = self.scope.projection(Namespace::Type, *view);
                }
                let new = self.scope.alias(kind.into(), binder(&source), &item)?;
                self.scope.publish(name.name.clone(), kind.into(), new)?;
                self.binary
                    .definitions
                    .push(export_definition(name, kind, index, ty));
            }
            Rule::comp_value => {
                let ty = self.valtype(child(&source, Rule::valtype)?)?;
                let value = child(&source, Rule::comp_val)?;
                let mut payload = Vec::new();
                if head(&value) == "binary" {
                    // This form deliberately needs no text-side type expansion;
                    // the native validator resolves aliases and checks the bytes.
                    literal(&ty, value, &mut payload, 0)?;
                } else {
                    literal(&self.shape(&ty)?, value, &mut payload, 0)?;
                }
                self.scope.add(Kind::Value, binder(&source))?;
                self.binary
                    .definitions
                    .push(Def::ValueLiteral { ty, payload });
            }
            Rule::comp_start => {
                let function = self
                    .scope
                    .index(Kind::Func, &child(&source, Rule::index)?)?;
                let mut args = Vec::new();
                let mut results = 0_u32;
                for p in source.into_inner() {
                    match p.as_rule() {
                        Rule::start_value => {
                            args.push(self.scope.index(Kind::Value, &child(&p, Rule::index)?)?)
                        }
                        Rule::start_result => {
                            self.scope.add(Kind::Value, binder(&p))?;
                            results = results.checked_add(1).ok_or("result count overflow")?;
                        }
                        Rule::index => (),
                        _ => return Err("unexpected start field".into()),
                    }
                }
                self.binary.definitions.push(Def::Start {
                    function,
                    args,
                    results,
                });
            }
            Rule::component => {
                let mut outers = self.outers.clone();
                outers.push(self.scope.clone());
                let (nested, interface) = produce_at(source.clone(), outers)?;
                let index = self.scope.add(Kind::Component, binder(&source))?;
                self.scope.define_shape(
                    Namespace::Component,
                    index,
                    SourceShape::Interface(interface),
                )?;
                self.binary
                    .definitions
                    .push(Def::NestedComponent(Arc::new(nested)));
            }
            Rule::comp_instance => {
                let expr = child(&source, Rule::instanceexpr)?;
                if let Some(instantiate) = expr
                    .clone()
                    .into_inner()
                    .find(|p| p.as_rule() == Rule::comp_instantiate)
                {
                    let component = self
                        .scope
                        .index(Kind::Component, &child(&instantiate, Rule::index)?)?;
                    let mut args = Vec::new();
                    for p in instantiate
                        .into_inner()
                        .filter(|p| p.as_rule() == Rule::instantiatearg)
                    {
                        let name = component_name(child(&p, Rule::string)?.as_str(), "argument")?;
                        let (kind, index) = if let Ok(target) = child(&p, Rule::externidx) {
                            self.item(target)?
                        } else {
                            (Kind::Instance, self.inline_instance(p.clone())?)
                        };
                        args.push(Item { name, kind, index });
                    }
                    self.binary
                        .definitions
                        .push(Def::ComponentInstance { component, args });
                    let projection = self
                        .scope
                        .projection(Namespace::Component, component)
                        .ok_or("component source projection is unavailable")?;
                    let SourceShape::Interface(interface) = &projection.shape else {
                        return Err("source component is not an interface".into());
                    };
                    self.define_instance(binder(&source), interface.clone())?;
                } else {
                    let mut exports = Vec::new();
                    for p in expr.into_inner() {
                        let name = external_name(&p)?;
                        let (kind, index) = self.item(child(&p, Rule::externidx)?)?;
                        exports.push(NamedItem { name, kind, index });
                    }
                    let interface = self.exported_items(&exports)?;
                    self.binary
                        .definitions
                        .push(Def::NamedInstanceFromExports(exports));
                    self.define_instance(binder(&source), interface)?;
                }
            }
            Rule::comp_alias => {
                if let Some(sort) = source
                    .clone()
                    .into_inner()
                    .find(|p| p.as_rule() == Rule::outeraliassort)
                {
                    let refs: Vec<_> = source
                        .clone()
                        .into_inner()
                        .filter(|p| p.as_rule() == Rule::index)
                        .collect();
                    let depth = if refs[0].as_str().starts_with('$') {
                        let label = &refs[0].as_str()[1..];
                        if self.scope.label.as_deref() == Some(label) {
                            0
                        } else {
                            self.outers
                                .iter()
                                .rev()
                                .position(|s| s.label.as_deref() == Some(label))
                                .ok_or("unknown outer scope")?
                                + 1
                        }
                    } else {
                        usize::try_from(parse_unsigned_const(refs[0].as_str())?)
                            .map_err(|_| "outer depth overflow")?
                    };
                    let scope = if depth == 0 {
                        &self.scope
                    } else {
                        self.outers
                            .get(
                                self.outers
                                    .len()
                                    .checked_sub(depth)
                                    .ok_or("outer scope out of bounds")?,
                            )
                            .ok_or("outer scope out of bounds")?
                    };
                    let k = kind(sort.as_str().trim())?;
                    let index = scope.index(k, &refs[1])?;
                    let item = scope.item(k.into(), index)?;
                    let target = match k {
                        Kind::Type => ComponentOuterAliasKind::Type,
                        Kind::Component => ComponentOuterAliasKind::Component,
                        Kind::Module => ComponentOuterAliasKind::CoreModule,
                        _ => return Err("invalid outer alias sort".into()),
                    };
                    self.scope.alias(k.into(), binder(&source), &item)?;
                    self.binary
                        .definitions
                        .push(Def::AliasOuter(ComponentOuterAlias {
                            kind: target,
                            count: u32::try_from(depth).map_err(|_| "outer depth overflow")?,
                            index,
                        }));
                } else {
                    let sort = child(&source, Rule::sort_c)?;
                    let k = kind(sort.as_str().trim())?;
                    let instance = self
                        .scope
                        .index(Kind::Instance, &child(&source, Rule::index)?)?;
                    let name =
                        component_name(child(&source, Rule::externnamelit)?.as_str(), "alias")?;
                    self.alias_export(instance, name, k, binder(&source))?;
                }
            }
            Rule::core_module_def => {
                // Preserve the existing core-module encoder; never give it a component.
                let mut text = String::from("(module");
                for p in source.clone().into_inner().filter(|p| {
                    matches!(p.as_rule(), Rule::id | Rule::module_field | Rule::rec_group)
                }) {
                    text.push(' ');
                    text.push_str(p.as_str());
                }
                text.push(')');
                let bytes = wat::parse_str(&text).map_err(|e| format!("core module: {e}"))?;
                self.scope.add(Kind::Module, binder(&source))?;
                self.binary.definitions.push(Def::CoreModule(bytes));
            }
            _ => {
                return Err(format!(
                    "primary value text producer does not yet support {:?}",
                    source.as_rule()
                ));
            }
        }
        Ok(())
    }
}

fn external_namespace(ty: &Ref) -> Namespace {
    match ty {
        Ref::Module(_) => Namespace::Module,
        Ref::Func(_) => Namespace::Func,
        Ref::Value(_) | Ref::ValueEq(_) => Namespace::Value,
        Ref::TypeEq(_) | Ref::TypeSubResource => Namespace::Type,
        Ref::Component(_) => Namespace::Component,
        Ref::Instance(_) => Namespace::Instance,
    }
}

fn export_definition(name: Name, kind: Kind, index: u32, ty: Option<Ref>) -> Def {
    if name.attributes.is_empty() {
        match (kind, &ty) {
            (Kind::Func, None) => {
                return Def::ExportFunc {
                    name: name.name,
                    index,
                    type_index: None,
                };
            }
            (Kind::Func, Some(Ref::Func(type_index))) => {
                return Def::ExportFunc {
                    name: name.name,
                    index,
                    type_index: Some(*type_index),
                };
            }
            (Kind::Type, None) => {
                return Def::ExportType {
                    name: name.name,
                    index,
                };
            }
            (Kind::Module, None) => {
                return Def::ExportCoreModule {
                    name: name.name,
                    index,
                };
            }
            (Kind::Instance, None) => {
                return Def::ExportInstance {
                    name: name.name,
                    index,
                };
            }
            _ => (),
        }
    }
    Def::NamedExport {
        name,
        kind,
        index,
        ty,
    }
}

fn interface_decl(definition: Def) -> Result<Decl, String> {
    Ok(match definition {
        Def::AliasOuter(ComponentOuterAlias {
            kind: ComponentOuterAliasKind::Type,
            count,
            index,
        }) => Decl::AliasOuterType { count, index },
        Def::AliasOuter(ComponentOuterAlias {
            kind: ComponentOuterAliasKind::CoreType,
            count,
            index,
        }) => Decl::AliasOuterCoreType { count, index },
        Def::CoreType(ty) => Decl::CoreType(ty),
        Def::AliasInstanceExport {
            instance,
            name,
            kind,
        } => Decl::AliasInstanceExport {
            instance,
            name,
            kind,
        },
        d @ (Def::ValueType(_)
        | Def::FuncType { .. }
        | Def::InterfaceType { .. }
        | Def::ResourceType { .. }) => Decl::Type(Box::new(d)),
        _ => return Err("invalid interface type declaration".into()),
    })
}

fn external_name(source: &Pair<'_, Rule>) -> Result<Name, String> {
    let name = component_name(
        child(source, Rule::externnamelit)?.as_str(),
        "external name",
    )?;
    let mut attributes = Vec::new();
    for p in source.clone().into_inner() {
        let p = if p.as_rule() == Rule::attribute {
            p.into_inner().next().ok_or("empty attribute")?
        } else {
            p
        };
        let attribute = match p.as_rule() {
            Rule::versionsuffix => Attribute::VersionSuffix(component_name(
                child(&p, Rule::string)?.as_str(),
                "version suffix",
            )?),
            Rule::implements_attr => Attribute::Implements(component_name(
                child(&p, Rule::externnamelit)?.as_str(),
                "implements",
            )?),
            Rule::externid_attr => Attribute::ExternalId(component_name(
                child(&p, Rule::string)?.as_str(),
                "external id",
            )?),
            _ => continue,
        };
        attributes.push(attribute);
    }
    Ok(Name { name, attributes })
}

fn produce_at(
    source: Pair<'_, Rule>,
    outers: Vec<SourceScope>,
) -> Result<(ComponentBinary, SourceInterface), String> {
    let mut producer = Producer {
        scope: SourceScope::new(&source),
        outers,
        ..Producer::default()
    };
    for p in source
        .into_inner()
        .filter(|p| p.as_rule() == Rule::component_definition)
    {
        producer.definition(p.into_inner().next().ok_or("empty component definition")?)?;
    }
    producer.scope.check()?;
    Ok((producer.binary, producer.scope.interface(true)))
}

pub(super) fn produce(source: Pair<'_, Rule>) -> Result<ComponentBinary, String> {
    produce_at(source, Vec::new()).map(|(binary, _)| binary)
}

fn uleb(mut value: u64, out: &mut Vec<u8>) {
    loop {
        let byte = (value & 127) as u8;
        value >>= 7;
        out.push(byte | if value == 0 { 0 } else { 128 });
        if value == 0 {
            break;
        }
    }
}

fn sleb(mut value: i64, out: &mut Vec<u8>) {
    loop {
        let byte = (value as u8) & 127;
        value >>= 7;
        let done = value == 0 && byte & 64 == 0 || value == -1 && byte & 64 != 0;
        out.push(byte | if done { 0 } else { 128 });
        if done {
            break;
        }
    }
}

fn count(n: usize, out: &mut Vec<u8>) -> Result<(), String> {
    uleb(
        u32::try_from(n).map_err(|_| "literal length overflow")? as u64,
        out,
    );
    Ok(())
}

fn literal(
    ty: &Val,
    source: Pair<'_, Rule>,
    out: &mut Vec<u8>,
    depth: usize,
) -> Result<(), String> {
    if depth > 512 {
        return Err("literal nesting limit exceeded".into());
    }
    let tag = head(&source);
    if tag == "binary" {
        for p in source.into_inner() {
            out.extend(unquote_to_bytes(p.as_str()));
        }
        return Ok(());
    }
    let text = source.as_str();
    let parts: Vec<_> = source
        .clone()
        .into_inner()
        .filter(|p| p.as_rule() == Rule::comp_val)
        .collect();
    let mismatch = || format!("literal does not match {ty:?}: {text}");
    match ty {
        Val::Prim(name) => match name.as_str() {
            "bool" if text == "true" || text == "false" => out.push(u8::from(text == "true")),
            "u8" | "u16" | "u32" | "u64" => {
                let value = parse_unsigned_const(text)?;
                let bits = name[1..].parse::<u32>().map_err(|_| mismatch())?;
                if bits < 64 && value >= 1_u64 << bits {
                    return Err(mismatch());
                }
                if bits == 8 {
                    out.push(value as u8);
                } else {
                    uleb(value, out);
                }
            }
            "s8" | "s16" | "s32" | "s64" => {
                let bits = name[1..].parse::<u32>().map_err(|_| mismatch())?;
                let clean = text.replace('_', "");
                let (negative, digits) = clean
                    .strip_prefix('-')
                    .map_or((false, clean.as_str()), |s| (true, s));
                let magnitude = parse_unsigned_const(digits)? as i128;
                let value = if negative { -magnitude } else { magnitude };
                if value < -(1_i128 << (bits - 1)) || value >= 1_i128 << (bits - 1) {
                    return Err(mismatch());
                }
                if bits == 8 {
                    out.push(value as u8);
                } else {
                    sleb(value as i64, out);
                }
            }
            "f32" | "f64" => {
                if text.contains("nan") && text != "nan" && text != "+nan" {
                    return Err("noncanonical literal NaN".into());
                }
                if find_rule(&source, Rule::float).is_none()
                    && find_rule(&source, Rule::integer).is_none()
                    && text != "nan"
                {
                    return Err(mismatch());
                }
                if super::float_literal_out_of_range(text, name == "f32") {
                    return Err("literal float out of range".into());
                }
                let expr = super::parse_float_at(text, name == "f32");
                let vybe_ast::ExprKind::Lit(vybe_ast::Literal::Float(value)) = expr.kind else {
                    return Err(mismatch());
                };
                if name == "f32" {
                    out.extend((value as f32).to_le_bytes());
                } else {
                    out.extend(value.to_le_bytes());
                }
            }
            "string" => {
                let p = child(&source, Rule::string)?;
                let bytes = unquote_to_bytes(p.as_str());
                std::str::from_utf8(&bytes).map_err(|_| "literal string is not UTF-8")?;
                count(bytes.len(), out)?;
                out.extend(bytes);
            }
            "char" => {
                let p = child(&source, Rule::char_lit)?;
                let raw = p.as_str();
                let bytes = unquote_to_bytes(&format!("\"{}\"", &raw[1..raw.len() - 1]));
                let value = std::str::from_utf8(&bytes).map_err(|_| "invalid literal char")?;
                if value.chars().count() != 1 {
                    return Err(mismatch());
                }
                out.extend(bytes);
            }
            _ => return Err(mismatch()),
        },
        Val::Record(fields) if tag == "record" => {
            if parts.len() != fields.len() {
                return Err(mismatch());
            }
            for ((_, ty), p) in fields.iter().zip(parts) {
                literal(ty, p, out, depth + 1)?;
            }
        }
        Val::Tuple(types) if tag == "tuple" => {
            if parts.len() != types.len() {
                return Err(mismatch());
            }
            for (ty, p) in types.iter().zip(parts) {
                literal(ty, p, out, depth + 1)?;
            }
        }
        Val::List(element) | Val::ListFixed(element, _) if tag == "list" => {
            // Fixed lists have no length prefix in the primary payload grammar.
            if let Val::ListFixed(_, n) = ty {
                if parts.len() != *n as usize {
                    return Err(mismatch());
                }
            } else {
                count(parts.len(), out)?;
            }
            for p in parts {
                literal(element, p, out, depth + 1)?;
            }
        }
        Val::Option(ty) => {
            if text == "none" {
                out.push(0);
            } else if tag == "some" && parts.len() == 1 {
                out.push(1);
                literal(ty, parts[0].clone(), out, depth + 1)?;
            } else {
                return Err(mismatch());
            }
        }
        Val::Result(ok, err) if tag == "ok" || tag == "error" => {
            let selected = if tag == "ok" {
                out.push(0);
                ok
            } else {
                out.push(1);
                err
            };
            match (selected, parts.as_slice()) {
                (None, []) => (),
                (Some(ty), [p]) => literal(ty, p.clone(), out, depth + 1)?,
                _ => return Err(mismatch()),
            }
        }
        Val::Variant(cases) if tag == "variant" => {
            let label = component_name(child(&source, Rule::labellit)?.as_str(), "case")?;
            let (index, (_, ty)) = cases
                .iter()
                .enumerate()
                .find(|(_, (n, _))| n == &label)
                .ok_or_else(mismatch)?;
            count(index, out)?;
            match (ty, parts.as_slice()) {
                (None, []) => (),
                (Some(ty), [p]) => literal(ty, p.clone(), out, depth + 1)?,
                _ => return Err(mismatch()),
            }
        }
        Val::Enum(names) if tag == "enum" => {
            let label = component_name(child(&source, Rule::labellit)?.as_str(), "enum")?;
            count(
                names
                    .iter()
                    .position(|n| n == &label)
                    .ok_or_else(mismatch)?,
                out,
            )?;
        }
        Val::Flags(names) if tag == "flags" => {
            let mut bytes = vec![0; names.len().div_ceil(8)];
            for p in source
                .into_inner()
                .filter(|p| p.as_rule() == Rule::labellit)
            {
                let label = component_name(p.as_str(), "flag")?;
                let i = names
                    .iter()
                    .position(|n| n == &label)
                    .ok_or("unknown literal flag")?;
                if bytes[i / 8] & (1 << (i % 8)) != 0 {
                    return Err("duplicate literal flag".into());
                }
                bytes[i / 8] |= 1 << (i % 8);
            }
            out.extend(bytes);
        }
        Val::Map(key, value) if tag == "map" => {
            let entries: Vec<_> = source
                .into_inner()
                .filter(|p| p.as_rule() == Rule::comp_entryval)
                .collect();
            count(entries.len(), out)?;
            for p in entries {
                let fields: Vec<_> = p.into_inner().collect();
                if fields.len() != 2 {
                    return Err("map entry needs two values".into());
                }
                literal(key, fields[0].clone(), out, depth + 1)?;
                literal(value, fields[1].clone(), out, depth + 1)?;
            }
        }
        _ => return Err(mismatch()),
    }
    Ok(())
}
