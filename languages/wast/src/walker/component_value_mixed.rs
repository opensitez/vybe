//! Core-only schema normalization and native mixed component definitions.
use super::*;
use vybe_ast::canon::{
    ComponentBinaryAsyncKind as AsyncKind, ComponentBinaryAsyncOperation as AsyncOp,
    ComponentBinaryCanonOption as Opt, ComponentCoreValType, ComponentCoreValueType,
    CoreExportKind as CoreKind, CoreInstanceExport,
};

fn core_kind(namespace: Namespace) -> Result<CoreKind, String> {
    match namespace {
        Namespace::CoreFunc => Ok(CoreKind::Func),
        Namespace::CoreTable => Ok(CoreKind::Table),
        Namespace::CoreMemory => Ok(CoreKind::Memory),
        Namespace::CoreGlobal => Ok(CoreKind::Global),
        Namespace::CoreTag => Ok(CoreKind::Tag),
        _ => Err("invalid core export namespace".into()),
    }
}

fn core_namespace(sort: &Pair<'_, Rule>) -> Result<Namespace, String> {
    let sort = if sort.as_rule() == Rule::core_sort {
        sort.clone()
    } else {
        child(sort, Rule::core_sort)?
    };
    match sort.as_str() {
        "func" => Ok(Namespace::CoreFunc),
        "table" => Ok(Namespace::CoreTable),
        "memory" => Ok(Namespace::CoreMemory),
        "global" => Ok(Namespace::CoreGlobal),
        "tag" => Ok(Namespace::CoreTag),
        "type" => Ok(Namespace::CoreType),
        "instance" => Ok(Namespace::CoreInstance),
        "module" => Ok(Namespace::Module),
        _ => Err("unknown core namespace".into()),
    }
}

impl Producer {
    fn canonical(
        &mut self,
        row: Pair<'_, Rule>,
        inline_bind: Option<String>,
    ) -> Result<(), String> {
        let bind =
            inline_bind.or_else(|| find_rule(&row, Rule::core_func_bind).and_then(|p| binder(&p)));
        let definition = match row.as_rule() {
            Rule::canon_lift => {
                let (ty, name) = self.external(child(&row, Rule::externtype)?)?;
                let core_func =
                    self.core_reference(child(&row, Rule::core_funcidx_ref)?, Namespace::CoreFunc)?;
                let options = self.options(row.clone())?;
                let Ref::Func(type_index) = ty else {
                    return Err("lift requires function type".into());
                };
                self.scope.add(Kind::Func, name)?;
                self.binary.definitions.push(Def::CanonLift {
                    core_func,
                    type_index,
                    options,
                });
                return Ok(());
            }
            Rule::canon_lower | Rule::canon_lower_inline => {
                let func = self.function_reference(child(&row, Rule::funcidx_ref)?)?;
                let options = self.options(row)?;
                Def::CanonLower { func, options }
            }
            _ => {
                let row = if row.as_rule() == Rule::canon_builtin {
                    row.into_inner().next().ok_or("missing canonical builtin")?
                } else {
                    row
                };
                let op = row
                    .clone()
                    .into_inner()
                    .next()
                    .ok_or("missing canonical operation")?;
                let options = self.options(row.clone())?;
                let type_index = |this: &mut Self| this.type_index(&child(&row, Rule::index)?);
                match op.as_str() {
                    "resource.new" => Def::CanonResourceNew {
                        type_index: type_index(self)?,
                    },
                    "resource.drop" => Def::CanonResourceDrop {
                        type_index: type_index(self)?,
                    },
                    "resource.rep" => Def::CanonResourceRep {
                        type_index: type_index(self)?,
                    },
                    "thread.index" => Def::CanonThreadIndex,
                    "thread.available-parallelism" => Def::CanonThreadAvailableParallelism {
                        shared: find_rule(&row, Rule::shared_kw).is_some(),
                    },
                    "thread.spawn-ref" => Def::CanonThreadSpawnRef {
                        shared: find_rule(&row, Rule::shared_kw).is_some(),
                        type_index: self.core_reference(
                            child(&row, Rule::core_typeidx_ref)?, Namespace::CoreType,
                        )?,
                    },
                    "thread.spawn-indirect" => Def::CanonThreadSpawnIndirect {
                        shared: find_rule(&row, Rule::shared_kw).is_some(),
                        type_index: self.core_reference(
                            child(&row, Rule::core_typeidx_ref)?, Namespace::CoreType,
                        )?,
                        table_index: self.core_reference(
                            child(&row, Rule::core_tableidx_ref)?, Namespace::CoreTable,
                        )?,
                    },
                    "thread.resume-later" => Def::CanonThreadResumeLater,
                    "thread.suspend"
                    | "thread.yield"
                    | "thread.suspend-then-resume"
                    | "thread.yield-then-resume"
                    | "thread.suspend-then-promote"
                    | "thread.yield-then-promote" => {
                        if find_rule(&row, Rule::cancellable_kw).is_some() {
                            return Err(
                                "cancellable canonical row requires native ABI representation"
                                    .into(),
                            );
                        }
                        match op.as_str() {
                            "thread.suspend" => Def::CanonThreadSuspend,
                            "thread.yield" => Def::CanonThreadYield,
                            "thread.suspend-then-resume" => Def::CanonThreadSuspendThenResume,
                            "thread.yield-then-resume" => Def::CanonThreadYieldThenResume,
                            "thread.suspend-then-promote" => Def::CanonThreadSuspendThenPromote,
                            _ => Def::CanonThreadYieldThenPromote,
                        }
                    }
                    "thread.new-indirect" => Def::CanonThreadNewIndirect {
                        type_index: self.core_reference(
                            child(&row, Rule::core_typeidx_ref)?,
                            Namespace::CoreType,
                        )?,
                        table_index: self.core_reference(
                            child(&row, Rule::core_tableidx_ref)?,
                            Namespace::CoreTable,
                        )?,
                    },
                    "waitable-set.new" => Def::CanonWaitableSetNew,
                    "waitable-set.drop" => Def::CanonWaitableSetDrop,
                    "waitable.join" => Def::CanonWaitableJoin,
                    "waitable-set.wait" | "waitable-set.poll" => {
                        if find_rule(&row, Rule::cancellable_kw).is_some() {
                            return Err(
                                "cancellable wait requires native ABI representation".into()
                            );
                        }
                        let memory = self.core_reference(
                            child(&row, Rule::core_memidx_ref)?,
                            Namespace::CoreMemory,
                        )?;
                        if op.as_str() == "waitable-set.wait" {
                            Def::CanonWaitableSetWait { memory }
                        } else {
                            Def::CanonWaitableSetPoll { memory }
                        }
                    }
                    "task.cancel" => Def::CanonTaskCancel,
                    "subtask.drop" => Def::CanonSubtaskDrop,
                    "subtask.cancel" => Def::CanonSubtaskCancel {
                        is_async: find_rule(&row, Rule::async_kw).is_some(),
                    },
                    "backpressure.inc" => Def::CanonBackpressureInc,
                    "backpressure.dec" => Def::CanonBackpressureDec,
                    "error-context.new" => Def::CanonErrorContextNew { options },
                    "error-context.debug-message" => Def::CanonErrorContextDebugMessage { options },
                    "error-context.drop" => Def::CanonErrorContextDrop,
                    "context.get" | "context.set" => {
                        let ty = match child(&row, Rule::val_type)?.as_str() {
                            "i32" => ComponentCoreValueType::Scalar(ComponentCoreValType::I32),
                            "i64" => ComponentCoreValueType::Scalar(ComponentCoreValType::I64),
                            _ => {
                                return Err(
                                    "context slot type requires native ABI representation".into()
                                );
                            }
                        };
                        let slot = u32::try_from(parse_unsigned_const(
                            child(&row, Rule::integer)?.as_str(),
                        )?)
                        .map_err(|_| "context slot overflow")?;
                        Def::CanonContext {
                            set: op.as_str() == "context.set",
                            ty,
                            slot,
                        }
                    }
                    "task.return" => Def::CanonTaskReturn {
                        result: row
                            .clone()
                            .into_inner()
                            .find(|p| p.as_rule() == Rule::func_result_c)
                            .map(|p| self.valtype(child(&p, Rule::valtype)?))
                            .transpose()?,
                        options,
                    },
                    name if name.starts_with("stream.") || name.starts_with("future.") => {
                        let (family, operation) =
                            name.split_once('.').ok_or("bad async operation")?;
                        let kind = if family == "stream" {
                            AsyncKind::Stream
                        } else {
                            AsyncKind::Future
                        };
                        let is_async = find_rule(&row, Rule::async_kw).is_some();
                        let operation = match operation {
                            "new" => AsyncOp::New,
                            "read" => AsyncOp::Read,
                            "write" => AsyncOp::Write,
                            "cancel-read" => AsyncOp::CancelRead { is_async },
                            "cancel-write" => AsyncOp::CancelWrite { is_async },
                            "drop-readable" => AsyncOp::DropReadable,
                            "drop-writable" => AsyncOp::DropWritable,
                            "forward" => AsyncOp::Forward,
                            _ => return Err("unknown async operation".into()),
                        };
                        Def::CanonAsync {
                            kind,
                            operation,
                            type_index: type_index(self)?,
                            options,
                        }
                    }
                    _ => {
                        return Err(format!(
                            "canonical operation {} requires native ABI representation",
                            op.as_str()
                        ));
                    }
                }
            }
        };
        self.binary.definitions.push(definition);
        self.scope.define(Namespace::CoreFunc, bind, None)?;
        Ok(())
    }

    pub(super) fn mixed_definition(&mut self, source: Pair<'_, Rule>) -> Result<bool, String> {
        match source.as_rule() {
            Rule::core_type_def | Rule::core_rec_group => {
                let names = if source.as_rule() == Rule::core_rec_group {
                    source
                        .clone()
                        .into_inner()
                        .filter(|p| p.as_rule() == Rule::core_member_type)
                        .map(|p| binder(&p))
                        .collect()
                } else {
                    vec![binder(&source)]
                };
                self.core_schema(source.as_str().into(), names)?;
            }
            Rule::comp_alias if find_rule(&source, Rule::core_outer_type_sort).is_some() => {
                self.core_schema(source.as_str().into(), vec![binder(&source)])?;
            }
            Rule::comp_alias
                if source
                    .clone()
                    .into_inner()
                    .any(|p| p.as_rule() == Rule::core_sort)
                    && find_rule(&source, Rule::outeraliassort).is_none() =>
            {
                let namespace = core_namespace(&child(&source, Rule::core_sort)?)?;
                let instance = self
                    .scope
                    .resolve(Namespace::CoreInstance, &child(&source, Rule::index)?)?;
                let name = component_name(child(&source, Rule::string)?.as_str(), "core alias")?;
                self.binary.definitions.push(Def::CoreAliasInstanceExport {
                    instance,
                    name,
                    kind: core_kind(namespace)?,
                });
                self.scope.define(namespace, binder(&source), None)?;
            }
            Rule::core_func_inline_alias => {
                let instance = self
                    .scope
                    .resolve(Namespace::CoreInstance, &child(&source, Rule::index)?)?;
                let name = component_name(child(&source, Rule::string)?.as_str(), "core alias")?;
                self.binary.definitions.push(Def::CoreAliasInstanceExport {
                    instance,
                    name,
                    kind: CoreKind::Func,
                });
                self.scope
                    .define(Namespace::CoreFunc, binder(&source), None)?;
            }
            Rule::core_instance_def => {
                let expr = child(&source, Rule::core_instanceexpr)?;
                if let Some(instantiate) = expr
                    .clone()
                    .into_inner()
                    .find(|p| p.as_rule() == Rule::core_instantiate)
                {
                    let module = self
                        .scope
                        .index(Kind::Module, &child(&instantiate, Rule::index)?)?;
                    let mut args = Vec::new();
                    for p in instantiate
                        .into_inner()
                        .filter(|p| p.as_rule() == Rule::core_instantiatearg)
                    {
                        let name =
                            component_name(child(&p, Rule::string)?.as_str(), "core argument")?;
                        let index = if p
                            .clone()
                            .into_inner()
                            .any(|p| p.as_rule() == Rule::core_inlineexport)
                        {
                            let exports = self.core_exports(p)?;
                            self.binary
                                .definitions
                                .push(Def::CoreInstanceFromExports(exports));
                            self.scope.define(Namespace::CoreInstance, None, None)?
                        } else {
                            self.scope
                                .resolve(Namespace::CoreInstance, &child(&p, Rule::index)?)?
                        };
                        args.push((name, index));
                    }
                    self.binary
                        .definitions
                        .push(Def::CoreInstance { module, args });
                } else {
                    let exports = self.core_exports(expr)?;
                    self.binary
                        .definitions
                        .push(Def::CoreInstanceFromExports(exports));
                }
                self.scope
                    .define(Namespace::CoreInstance, binder(&source), None)?;
            }
            Rule::canon => {
                let row = child(&source, Rule::canon_body)?
                    .into_inner()
                    .next()
                    .ok_or("empty canonical body")?;
                self.canonical(row, None)?;
            }
            Rule::core_func_inline_lower => {
                self.canonical(child(&source, Rule::canon_lower_inline)?, binder(&source))?
            }
            Rule::core_func_inline_builtin => {
                self.canonical(child(&source, Rule::canon_builtin)?, binder(&source))?
            }
            Rule::comp_func => {
                let mut alias_item = None;
                if let Ok(lift) = child(&source, Rule::comp_inline_lift) {
                    let type_index = self.function_type(child(&source, Rule::comp_func_type)?)?;
                    let core_func = self.core_reference(
                        child(&lift, Rule::core_funcidx_ref)?,
                        Namespace::CoreFunc,
                    )?;
                    let options = self.options(lift)?;
                    self.binary.definitions.push(Def::CanonLift {
                        core_func,
                        type_index,
                        options,
                    });
                } else {
                    if child(&source, Rule::comp_func_type)?
                        .into_inner()
                        .next()
                        .is_some()
                    {
                        return Err(
                            "inline function alias: type ascription is not checked yet".into()
                        );
                    }
                    let alias = child(&source, Rule::comp_inline_alias)?;
                    let instance = self
                        .scope
                        .index(Kind::Instance, &child(&alias, Rule::index)?)?;
                    let name = component_name(
                        child(&alias, Rule::externnamelit)?.as_str(),
                        "function alias",
                    )?;
                    let item = self.scope.export(Namespace::Instance, instance, &name)?;
                    if item.namespace != Namespace::Func {
                        return Err("source export has the wrong namespace".into());
                    }
                    alias_item = Some(item);
                    self.binary.definitions.push(Def::AliasInstanceExport {
                        instance,
                        name,
                        kind: Kind::Func,
                    });
                }
                let index = if let Some(item) = alias_item {
                    self.scope.alias(Namespace::Func, binder(&source), &item)?
                } else {
                    self.scope.add(Kind::Func, binder(&source))?
                };
                for p in source
                    .into_inner()
                    .filter(|p| p.as_rule() == Rule::comp_func_export)
                {
                    let name = external_name(&p)?;
                    self.scope
                        .publish(name.name.clone(), Namespace::Func, index)?;
                    self.binary
                        .definitions
                        .push(export_definition(name, Kind::Func, index, None));
                    self.scope.define(
                        Namespace::Func,
                        None,
                        Some(self.scope.origin(Namespace::Func, index)?),
                    )?;
                }
            }
            _ => return Ok(false),
        }
        Ok(true)
    }

    // Only core schema declarations and their lexical spine enter this resolver.
    // Component values, functions, exports and canonical rows never enter it.
    fn core_schema(&mut self, text: String, names: Vec<Option<String>>) -> Result<u32, String> {
        let mut shell = String::new();
        for scope in self.outers.iter().chain(std::iter::once(&self.scope)) {
            shell.push_str("(component ");
            if let Some(label) = &scope.label {
                shell.push('$');
                shell.push_str(label);
                shell.push(' ');
            }
            for declaration in &scope.core_context {
                shell.push_str(declaration);
                shell.push('\n');
            }
        }
        shell.push_str(&text);
        for _ in 0..=self.outers.len() {
            shell.push(')');
        }
        let bytes = wat::parse_str(&shell).map_err(|e| format!("core schema subtree: {e}"))?;
        let mut binary = vybe_platform_wasm::read_component_binary(&bytes)?;
        for _ in 0..self.outers.len() {
            binary = binary
                .definitions
                .iter()
                .rev()
                .find_map(|d| {
                    if let Def::NestedComponent(child) = d {
                        Some((**child).clone())
                    } else {
                        None
                    }
                })
                .ok_or("core schema lexical scope lost")?;
        }
        let definitions: Vec<_> = std::mem::take(&mut binary.definitions)
            .into_iter()
            .filter(|d| matches!(d, Def::CoreType(_)) || d.normalized_outer_alias().is_some())
            .collect();
        let added = definitions
            .get(self.scope.core_definitions..)
            .ok_or("core schema declaration prefix changed")?;
        let mut pending = Vec::new();
        for definition in added {
            let count = if let Some(ty) = definition.normalized_core_type() {
                ty.type_count().ok_or("core type count overflow")?
            } else {
                1
            };
            for member in 0..count {
                pending.push((definition, member));
            }
        }
        if pending.len() < names.len() {
            return Err("core schema source binder count mismatch".into());
        }
        let bind_start = pending.len() - names.len();
        let first_named = self.scope.counts[Namespace::CoreType.slot()] + bind_start as u32;
        for (position, (definition, _)) in pending.into_iter().enumerate() {
            let origin = if let Some(alias) = definition.normalized_outer_alias() {
                let scope = if alias.count == 0 {
                    &self.scope
                } else {
                    self.outers
                        .get(
                            self.outers
                                .len()
                                .checked_sub(alias.count as usize)
                                .ok_or("core outer scope out of bounds")?,
                        )
                        .ok_or("core outer scope out of bounds")?
                };
                Some(scope.origin(Namespace::CoreType, alias.index)?)
            } else {
                None
            };
            self.scope.define(
                Namespace::CoreType,
                position
                    .checked_sub(bind_start)
                    .and_then(|i| names.get(i).cloned())
                    .flatten(),
                origin,
            )?;
        }
        self.binary.definitions.extend_from_slice(added);
        self.scope.core_definitions = definitions.len();
        self.scope.core_context.push(text);
        Ok(first_named)
    }

    pub(super) fn module_external(
        &mut self,
        source: &Pair<'_, Rule>,
    ) -> Result<Option<(Ref, Option<String>)>, String> {
        if let Some(module) = source
            .clone()
            .into_inner()
            .find(|p| p.as_rule() == Rule::core_module_type_ref)
        {
            let index = self
                .scope
                .resolve(Namespace::CoreType, &child(&module, Rule::index)?)?;
            return Ok(Some((Ref::Module(index), binder(&module))));
        }
        if source
            .clone()
            .into_inner()
            .any(|p| p.as_rule() == Rule::core_moduledecl)
        {
            let mut text = String::from("(core type (module ");
            for p in source
                .clone()
                .into_inner()
                .filter(|p| p.as_rule() == Rule::core_moduledecl)
            {
                text.push_str(p.as_str());
                text.push('\n');
            }
            text.push_str("))");
            let index = self.core_schema(text, vec![None])?;
            return Ok(Some((Ref::Module(index), binder(source))));
        }
        if source.clone().into_inner().all(|p| p.as_rule() == Rule::id) && head(source) == "core" {
            let index = self.core_schema("(core type (module))".into(), vec![None])?;
            return Ok(Some((Ref::Module(index), binder(source))));
        }
        Ok(None)
    }

    pub(super) fn core_reference(
        &mut self,
        source: Pair<'_, Rule>,
        namespace: Namespace,
    ) -> Result<u32, String> {
        let index = child(&source, Rule::index)?;
        let paths: Vec<_> = source
            .into_inner()
            .filter(|p| p.as_rule() == Rule::string)
            .collect();
        if paths.is_empty() {
            return self.scope.resolve(namespace, &index);
        }
        if paths.len() != 1 {
            return Err("core export path must have one field".into());
        }
        let instance = self.scope.resolve(Namespace::CoreInstance, &index)?;
        self.binary.definitions.push(Def::CoreAliasInstanceExport {
            instance,
            name: component_name(paths[0].as_str(), "core export")?,
            kind: core_kind(namespace)?,
        });
        self.scope.define(namespace, None, None)
    }

    fn core_exports(&mut self, source: Pair<'_, Rule>) -> Result<Vec<CoreInstanceExport>, String> {
        let mut exports = Vec::new();
        for p in source
            .into_inner()
            .filter(|p| p.as_rule() == Rule::core_inlineexport)
        {
            let name = component_name(child(&p, Rule::string)?.as_str(), "core export")?;
            let target = child(&p, Rule::core_externidx)?;
            let namespace = core_namespace(&target)?;
            let index = self.core_reference(target, namespace)?;
            exports.push(CoreInstanceExport {
                name,
                kind: core_kind(namespace)?,
                index,
            });
        }
        Ok(exports)
    }

    pub(super) fn inline_instance(&mut self, source: Pair<'_, Rule>) -> Result<u32, String> {
        let mut exports = Vec::new();
        for p in source
            .into_inner()
            .filter(|p| p.as_rule() == Rule::inlineexport)
        {
            let name = external_name(&p)?;
            let (kind, index) = self.item(child(&p, Rule::externidx)?)?;
            exports.push(NamedItem { name, kind, index });
        }
        let interface = self.exported_items(&exports)?;
        self.binary
            .definitions
            .push(Def::NamedInstanceFromExports(exports));
        self.define_instance(None, interface)
    }

    fn options(&mut self, source: Pair<'_, Rule>) -> Result<Vec<Opt>, String> {
        let mut options = Vec::new();
        for p in source
            .into_inner()
            .filter(|p| p.as_rule() == Rule::canonopt)
        {
            let opt = if let Some(encoding) = p
                .clone()
                .into_inner()
                .find(|p| p.as_rule() == Rule::string_encoding_opt)
            {
                match encoding.as_str() {
                    "string-encoding=utf8" => Opt::Utf8,
                    "string-encoding=utf16" => Opt::Utf16,
                    "string-encoding=latin1+utf16" => Opt::CompactUtf16,
                    _ => return Err("unknown canonical string encoding".into()),
                }
            } else if find_rule(&p, Rule::async_kw).is_some() {
                Opt::Async
            } else if let Some(reference) = find_rule(&p, Rule::core_memidx_ref) {
                Opt::Memory(self.core_reference(reference, Namespace::CoreMemory)?)
            } else {
                let index =
                    self.core_reference(child(&p, Rule::core_funcidx_ref)?, Namespace::CoreFunc)?;
                match head(&p) {
                    "realloc" => Opt::Realloc(index),
                    "post-return" => Opt::PostReturn(index),
                    "callback" => Opt::Callback(index),
                    _ => return Err("unknown canonical option".into()),
                }
            };
            options.push(opt);
        }
        Ok(options)
    }

    fn function_type(&mut self, source: Pair<'_, Rule>) -> Result<u32, String> {
        if let Some(reference) = source
            .clone()
            .into_inner()
            .find(|p| p.as_rule() == Rule::index)
        {
            return self.type_index(&reference);
        }
        let mut params = Vec::new();
        let mut result = None;
        let mut is_async = false;
        for p in source.into_inner() {
            match p.as_rule() {
                Rule::func_param_c => params.push((
                    component_name(child(&p, Rule::labellit)?.as_str(), "parameter")?,
                    self.valtype(child(&p, Rule::valtype)?)?,
                )),
                Rule::func_result_c => result = Some(self.valtype(child(&p, Rule::valtype)?)?),
                Rule::async_kw => is_async = true,
                _ => (),
            }
        }
        self.add_type(
            Def::FuncType {
                params,
                result,
                is_async,
            },
            None,
        )
    }

    fn function_reference(&mut self, source: Pair<'_, Rule>) -> Result<u32, String> {
        let reference = child(&source, Rule::index)?;
        let paths: Vec<_> = source
            .into_inner()
            .filter(|p| p.as_rule() == Rule::string)
            .collect();
        if paths.is_empty() {
            return self.scope.index(Kind::Func, &reference);
        }
        let mut instance = self.scope.index(Kind::Instance, &reference)?;
        for (position, path) in paths.iter().enumerate() {
            let kind = if position + 1 == paths.len() {
                Kind::Func
            } else {
                Kind::Instance
            };
            let index = self.alias_export(
                instance,
                component_name(path.as_str(), "component export path")?,
                kind,
                None,
            )?;
            if kind == Kind::Func {
                return Ok(index);
            }
            instance = index;
        }
        Err("empty function export path".into())
    }
}
