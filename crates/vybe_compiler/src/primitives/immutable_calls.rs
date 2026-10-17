//! Compiler-owned source binding receipts. Names index the unit's binder;
//! only the actual declaration producer can publish a native chunk identity.
use super::*;

#[cfg(test)]
#[path = "immutable_expression_inventory_tests.rs"]
mod expression_inventory_tests;

#[derive(Clone, Copy)]
struct CompletedEntry {
    chunk: usize,
    arity: u8,
    takes_receiver: bool,
    results: u8,
}

struct Declaration {
    statement: usize,
    params: usize,
    body: usize,
    params_len: usize,
    body_len: usize,
    completed: Option<CompletedEntry>,
}

#[derive(Default)]
pub(super) struct Bindings {
    by_binding: HashMap<String, usize>,
    declarations: HashMap<usize, Declaration>,
    claimed: HashMap<String, Option<usize>>,
    excluded: HashSet<String>,
    active: Option<usize>,
}

// Enumerate only containers whose complete declaration/write surface is known.
// Unknown containers miss the optimization; source syntax is never rejected.
enum BindingNode<'a> {
    Statement(&'a Statement),
    Expression(&'a Expression),
}

#[derive(Default)]
struct BindingWalk {
    statements: HashSet<usize>,
    expressions: HashSet<usize>,
    #[cfg(test)]
    expression_callbacks: usize,
}

fn visit_binding_expression(expression: &Expression, walk: &mut BindingWalk, visit: &mut impl FnMut(BindingNode<'_>)) -> bool {
    if walk.expressions.contains(&(expression as *const Expression as usize)) { return true; }
    let mut known = true;
    expression.walk_exprs(&mut |expression| {
        #[cfg(test)]
        { walk.expression_callbacks += 1; }
        known &= visit_binding_expression_node(expression, walk, visit);
    });
    known
}

fn visit_binding_expression_node(expression: &Expression, walk: &mut BindingWalk, visit: &mut impl FnMut(BindingNode<'_>)) -> bool {
    if !walk.expressions.insert(expression as *const Expression as usize) { return true; }
    visit(BindingNode::Expression(expression));
    match &expression.kind {
        ExprKind::Lambda { params, body, .. } => {
            if !params.iter().all(|param| param.default.as_ref()
                .is_none_or(|value| visit_binding_expression(value, walk, visit))) { return false; }
            match body {
                // walk_exprs already visits the expression body.
                LambdaBody::Expr(_) => true,
                LambdaBody::Block(body) => body.iter().all(|statement| visit_binding_statement_nodes(statement, walk, visit)),
            }
        }
        ExprKind::FunctionExpr(declaration) => match &declaration.kind {
            // The actual publisher invents a chunk-dependent name when empty.
            StmtKind::FunctionDecl { name, .. } if !name.is_empty() => visit_binding_statement_nodes(declaration, walk, visit),
            _ => false,
        },
        ExprKind::ClassExpr { name, parent, members, .. } => {
            // Anonymous names and dynamic parent globals depend on emission
            // state. Do not predict them from a source binding inventory.
            name.as_ref().is_some_and(|name| !name.is_empty()) && parent.is_none()
                && visit_binding_members(members, false, walk, visit)
        }
        ExprKind::Object(properties) => properties.iter().all(|property| match property {
            // The AST expression visitor skips these statement-valued bodies.
            ObjectProperty::Method { value, .. } | ObjectProperty::Accessor { value, .. } =>
                visit_binding_method(value, false, walk, visit),
            ObjectProperty::KeyValue { .. } | ObjectProperty::Computed { .. }
            | ObjectProperty::Spread(_) | ObjectProperty::Shorthand(_) => true,
        }),
        ExprKind::ArrayMap { array, params, body } =>
            visit_binding_expression(array, walk, visit)
                && params.iter().all(|param| param.default.as_ref()
                    .is_none_or(|value| visit_binding_expression(value, walk, visit)))
                && visit_binding_expression(body, walk, visit),
        ExprKind::ArrayTransform { args, .. } => args.iter()
            .all(|value| visit_binding_expression(value, walk, visit)),
        ExprKind::WasmCallWithTag { callee, args, .. } =>
            visit_binding_expression(callee, walk, visit)
                && args.iter().all(|value| visit_binding_expression(value, walk, visit)),
        ExprKind::RefOf(place) => match place.as_ref() {
            PlaceExpr::Ident(_) => true,
            PlaceExpr::Deref(value) | PlaceExpr::Member { object: value, .. } =>
                visit_binding_expression(value, walk, visit),
            PlaceExpr::Index { object, index, .. } => visit_binding_expression(object, walk, visit)
                && visit_binding_expression(index, walk, visit),
        },
        ExprKind::Destructure(pattern) => match pattern {
            DestructurePattern::Object(properties) =>
                inventory_object_pattern_defaults(properties, walk, visit),
            DestructurePattern::Array(elements) =>
                inventory_array_pattern_defaults(elements, walk, visit),
        },
        ExprKind::CallableRef { adapter, .. } => adapter.as_ref().is_none_or(|adapter| match adapter {
            CallableAdapter::Expr { params, .. } => params.iter().all(|param| param.default.as_ref()
                .is_none_or(|value| visit_binding_expression(value, walk, visit))),
        }),
        // These variants are leaves or have all expression children covered by
        // the AST visitor. Keep this exhaustive: new variants require an audit.
        ExprKind::Lit(_) | ExprKind::Ident(_) | ExprKind::This | ExprKind::Super
        | ExprKind::GlobalNamespace | ExprKind::Binary { .. } | ExprKind::Unary { .. }
        | ExprKind::RefLoad(_) | ExprKind::Ternary { .. } | ExprKind::Member { .. }
        | ExprKind::Index { .. } | ExprKind::Call { .. } | ExprKind::New { .. }
        | ExprKind::Assign { .. } | ExprKind::Array(_) | ExprKind::Tuple(_)
        | ExprKind::NamedTuple { .. } | ExprKind::Set(_) | ExprKind::Map(_)
        | ExprKind::Zip { .. } | ExprKind::Interpolation(_) | ExprKind::IsType { .. }
        | ExprKind::Cast { .. } | ExprKind::TypeOf(_) | ExprKind::DefaultOf(_)
        | ExprKind::NullCoalesce { .. } | ExprKind::Spread(_) | ExprKind::Async(_)
        | ExprKind::Chan(_) | ExprKind::Atomic(_) | ExprKind::Await(_) | ExprKind::Yield(_)
        | ExprKind::YieldFrom(_) | ExprKind::FuncRef(_) | ExprKind::SuperCall { .. }
        | ExprKind::Comprehension { .. } | ExprKind::Slice { .. } | ExprKind::Walrus { .. }
        | ExprKind::Void(_) | ExprKind::Delete(_) | ExprKind::Sequence(_)
        | ExprKind::Proxy { .. } | ExprKind::Range { .. } | ExprKind::StaticAccess { .. }
        | ExprKind::Match { .. } => true,
    }
}

fn visit_binding_tree(statement: &Statement, walk: &mut BindingWalk, visit: &mut impl FnMut(BindingNode<'_>)) -> bool {
    // One recursive expression walk per source root, not per nested statement.
    // Its expression-contained bodies are inventoried without rewalking them.
    let mut expressions_known = true;
    statement.walk_exprs(&mut |expression| {
        #[cfg(test)]
        { walk.expression_callbacks += 1; }
        expressions_known &= visit_binding_expression_node(expression, walk, visit);
    });
    expressions_known && visit_binding_statement_nodes(statement, walk, visit)
}

fn visit_binding_statement_nodes(statement: &Statement, walk: &mut BindingWalk, visit: &mut impl FnMut(BindingNode<'_>)) -> bool {
    if !walk.statements.insert(statement as *const Statement as usize) { return true; }
    visit(BindingNode::Statement(statement));
    let mut body = |body: &[Statement]| body.iter().all(|statement| visit_binding_statement_nodes(statement, walk, visit));
    match &statement.kind {
        StmtKind::FunctionDecl { body: statements, .. } | StmtKind::Block(statements)
        | StmtKind::While { body: statements, else_body: None, .. }
        | StmtKind::DoWhile { body: statements, .. } => body(statements),
        StmtKind::While { body: statements, else_body, .. }
        | StmtKind::ForIn { body: statements, else_body, .. } =>
            body(statements) && else_body.as_ref().is_none_or(|statements| body(statements)),
        StmtKind::For { init, body: statements, .. } => {
            if !body(statements) { return false; }
            init.as_ref().is_none_or(|statement| visit_binding_statement_nodes(statement, walk, visit))
        }
        StmtKind::If { then_body, elifs, else_body, .. } => body(then_body)
            && elifs.iter().all(|(_, statements)| body(statements))
            && else_body.as_ref().is_none_or(|statements| body(statements)),
        StmtKind::Try { body: statements, catches, else_body, finally } => body(statements)
            && catches.iter().all(|catch| body(&catch.body))
            && else_body.as_ref().is_none_or(|statements| body(statements))
            && finally.as_ref().is_none_or(|statements| body(statements)),
        StmtKind::Switch { cases, default, .. } => {
            if !cases.iter().all(|case| body(&case.body))
                || !default.as_ref().is_none_or(|statements| body(statements)) { return false; }
            cases.iter().all(|case| case.conditions.iter().all(|condition| match condition {
                CaseCondition::Value(value) | CaseCondition::Comparison { expr: value, .. } =>
                    visit_binding_expression(value, walk, visit),
                CaseCondition::Range { from, to } => visit_binding_expression(from, walk, visit)
                    && visit_binding_expression(to, walk, visit),
            }))
        }
        StmtKind::Labeled { body, .. } => visit_binding_statement_nodes(body, walk, visit),
        StmtKind::ClassDecl { members, decorators, .. } | StmtKind::StructDecl { members, decorators, .. } =>
            decorators.iter().all(|value| visit_binding_expression(value, walk, visit))
                && visit_binding_members(members, true, walk, visit),
        // Current AST visitors cover status already. The expression-root guard
        // avoids a second walk while retaining explicit coverage of this slot.
        StmtKind::Exit { status } => status.as_ref()
            .is_none_or(|value| visit_binding_expression(value, walk, visit)),
        StmtKind::Assert { test, msg } => visit_binding_expression(test, walk, visit)
            && msg.as_ref().is_none_or(|value| visit_binding_expression(value, walk, visit)),
        StmtKind::DataSegment { offset, .. } => offset.as_ref()
            .is_none_or(|value| visit_binding_expression(value, walk, visit)),
        StmtKind::MemoryDecl { .. } | StmtKind::GoTo(_) => true,
        StmtKind::Directive { set, .. } => binding_inert_directive(set),
        StmtKind::VarDecl { declarations, .. } => declarations.iter().all(|declaration|
            inventory_binding_pattern_defaults(&declaration.pattern, walk, visit)
                && declaration.array_bounds.as_ref().is_none_or(|bounds| bounds.iter()
                    .all(|value| visit_binding_expression(value, walk, visit)))
                && declaration.init.as_ref().is_none_or(|value| visit_binding_expression(value, walk, visit))),
        StmtKind::Break(BreakTarget::Value(value)) => visit_binding_expression(value, walk, visit),
        StmtKind::Expr(_) | StmtKind::Return(_)
        | StmtKind::Assign { .. } | StmtKind::CompoundAssign { .. }
        | StmtKind::Throw { .. } | StmtKind::Break(_) | StmtKind::Continue(_)
        | StmtKind::Empty => true,
        _ => false,
    }
}

fn inventory_binding_pattern_defaults(pattern: &BindingPattern, walk: &mut BindingWalk,
    visit: &mut impl FnMut(BindingNode<'_>)) -> bool {
    match pattern {
        BindingPattern::Ident(_) => true,
        BindingPattern::Object(properties) => inventory_object_pattern_defaults(properties, walk, visit),
        BindingPattern::Array(elements) => inventory_array_pattern_defaults(elements, walk, visit),
    }
}

fn inventory_object_pattern_defaults(properties: &[ObjectPatternProp], walk: &mut BindingWalk,
    visit: &mut impl FnMut(BindingNode<'_>)) -> bool {
    properties.iter().all(|property| property.default.as_ref()
        .is_none_or(|value| visit_binding_expression(value, walk, visit))
        && property.value.as_ref().is_none_or(|pattern| inventory_binding_pattern_defaults(pattern, walk, visit)))
}

fn inventory_array_pattern_defaults(elements: &[ArrayPatternElem], walk: &mut BindingWalk,
    visit: &mut impl FnMut(BindingNode<'_>)) -> bool {
    elements.iter().all(|element| match element {
        ArrayPatternElem::Pattern(pattern, default) => inventory_binding_pattern_defaults(pattern, walk, visit)
            && default.as_ref().is_none_or(|value| visit_binding_expression(value, walk, visit)),
        ArrayPatternElem::Rest(_) | ArrayPatternElem::Hole => true,
    })
}

fn binding_inert_directive(set: &Directives) -> bool {
    // Numeric value transport does not change source publication, namespace
    // lookup or callable entry ownership. Reject every unexamined policy field.
    set == &Directives { wide_integer_transport: set.wide_integer_transport,
        ..Default::default() }
}

fn visit_binding_member_body(body: &[Statement], expressions_walked: bool, walk: &mut BindingWalk,
    visit: &mut impl FnMut(BindingNode<'_>)) -> bool {
    body.iter().all(|statement| if expressions_walked {
        visit_binding_statement_nodes(statement, walk, visit)
    } else { visit_binding_tree(statement, walk, visit) })
}

fn visit_binding_method(declaration: &Statement, expressions_walked: bool, walk: &mut BindingWalk,
    visit: &mut impl FnMut(BindingNode<'_>)) -> bool {
    let StmtKind::FunctionDecl { params, body, .. } = &declaration.kind else { return false; };
    // The method's own name is a property, not a source function publication.
    (expressions_walked || params.iter().all(|param| param.default.as_ref()
        .is_none_or(|value| visit_binding_expression(value, walk, visit))))
        && visit_binding_member_body(body, expressions_walked, walk, visit)
}

fn visit_binding_members(members: &[ClassMember], expressions_walked: bool, walk: &mut BindingWalk, visit: &mut impl FnMut(BindingNode<'_>)) -> bool {
    for member in members {
        let known = match member {
            // Methods themselves publish member slots, NOT function globals.
            // Only declarations inside their bodies enter this inventory.
            ClassMember::Method(declaration) =>
                visit_binding_method(declaration, expressions_walked, walk, visit),
            ClassMember::Constructor { params, base_args, body, .. } =>
                params.iter().all(|param| param.default.as_ref()
                    .is_none_or(|value| visit_binding_expression(value, walk, visit)))
                && base_args.as_ref().is_none_or(|args| args.iter()
                    .all(|value| visit_binding_expression(value, walk, visit)))
                && visit_binding_member_body(body, expressions_walked, walk, visit),
            ClassMember::Property { getter, setter, .. } =>
                getter.as_ref().is_none_or(|body|visit_binding_member_body(body, expressions_walked, walk, visit))
                && setter.as_ref().is_none_or(|setter|
                    setter.param.default.as_ref().is_none_or(|value| visit_binding_expression(value, walk, visit))
                        && visit_binding_member_body(&setter.body, expressions_walked, walk, visit)),
            ClassMember::NestedType(statement) => if expressions_walked {
                visit_binding_statement_nodes(statement,walk,visit)
            } else { visit_binding_tree(statement,walk,visit) },
            ClassMember::Field { init, array_bounds, .. } =>
                (expressions_walked || init.as_ref().is_none_or(|value| visit_binding_expression(value, walk, visit)))
                && array_bounds.as_ref().is_none_or(|bounds| bounds.iter()
                    .all(|value| visit_binding_expression(value, walk, visit))),
            ClassMember::Const { value, .. } => expressions_walked || visit_binding_expression(value, walk, visit),
            ClassMember::Event { params, .. } => params.iter().all(|param| param.default.as_ref()
                .is_none_or(|value| visit_binding_expression(value, walk, visit))),
            // Imported augmentation bodies need their own publication inventory.
            ClassMember::Augment(_) => false,
        };
        if !known {return false;}
    }
    true
}

impl Bindings {
    fn prepare(body: &[Statement], canon: impl Fn(&str) -> String) -> Self {
        Self::prepare_with_aliases(body, canon, |_| Vec::new())
    }

    fn prepare_with_aliases(body: &[Statement], canon: impl Fn(&str) -> String,
        publication_names: impl Fn(&str) -> Vec<String>) -> Self {
        let mut bindings = Self::default();
        let mut seen = HashSet::new();
        let mut excluded = HashSet::new();
        for (ordinal, statement) in body.iter().enumerate() {
            match &statement.kind {
                StmtKind::FunctionDecl { name, params, body, handles, is_async, is_generator, .. } => {
                    let binding = canon(name);
                    if !seen.insert(binding.clone()) { excluded.insert(binding.clone()); }
                    if *is_async || *is_generator || !handles.is_empty()
                        || params.iter().any(|param| param.is_rest || param.is_kwargs
                            || param.default.is_some() || param.is_optional
                            || param.pass_by != PassBy::Value)
                    { continue; }
                    bindings.by_binding.insert(binding, ordinal);
                    bindings.declarations.insert(ordinal, Declaration {
                        statement: statement as *const Statement as usize,
                        params: params.as_ptr() as usize,
                        body: body.as_ptr() as usize,
                        params_len: params.len(), body_len: body.len(), completed: None,
                    });
                }
                _ => {
                    let mut names = HashSet::new();
                    collect_declared_names(std::slice::from_ref(statement), &mut names);
                    excluded.extend(names.iter().map(|name| canon(name)));
                }
            }
            // Writes are excluded by the single publication/expression walk
            // below, including writes inside expression-contained bodies.
        }
        // Preclaim the SAME aliases the publisher will emit. This inventory
        // does not publish receipts. It prevents a later conditional producer
        // or lossy alias encoding from invalidating an already-emitted edge.
        let mut walk = BindingWalk::default();
        for statement in body {
            if !visit_binding_tree(statement, &mut walk, &mut |node| {
                match node {
                BindingNode::Statement(statement) => match &statement.kind {
                    StmtKind::FunctionDecl { name, .. } => {
                        let source = canon(name);
                        let owner = bindings.by_binding.get(&source).copied().filter(|id|
                            bindings.declarations[id].statement == statement as *const Statement as usize);
                        bindings.claim(source, owner);
                        for alias in publication_names(name) { bindings.claim(canon(&alias), owner); }
                    }
                    StmtKind::Assign { targets, .. } => for target in targets {
                        target.walk_exprs(&mut |expression| if let ExprKind::Ident(name)=&expression.kind {
                            excluded.insert(canon(name));
                        });
                    },
                    StmtKind::CompoundAssign { target, .. } => target.walk_exprs(&mut |expression|
                        if let ExprKind::Ident(name)=&expression.kind { excluded.insert(canon(name)); }),
                    StmtKind::ForIn {var,key,..} => {
                        excluded.insert(canon(var));
                        if let Some(key)=key { excluded.insert(canon(key)); }
                    }
                    StmtKind::ClassDecl {name,..} | StmtKind::StructDecl {name,..} => {
                        excluded.insert(canon(name));
                    }
                    _ => {},
                },
                BindingNode::Expression(expression) => {
                    if let ExprKind::Lambda { captures, .. } = &expression.kind {
                        for capture in captures {
                            let (by_ref, name) = Compiler::split_explicit_capture(capture);
                            if by_ref { excluded.insert(canon(name)); }
                        }
                    }
                    if let ExprKind::ClassExpr { name: Some(name), .. } = &expression.kind {
                        // ClassExpr emits a global class binding, not function aliases.
                        bindings.claim(canon(name), None);
                    }
                    if let ExprKind::Assign { target, .. } | ExprKind::Walrus { target, .. }=&expression.kind {
                        target.walk_exprs(&mut |expression| if let ExprKind::Ident(name)=&expression.kind {
                            excluded.insert(canon(name));
                        });
                    }
                    if let ExprKind::Unary { op: UnaryOp::PreInc | UnaryOp::PostInc
                        | UnaryOp::PreDec | UnaryOp::PostDec | UnaryOp::AddrOf, expr }=&expression.kind {
                        expr.walk_exprs(&mut |expression| if let ExprKind::Ident(name)=&expression.kind {
                            excluded.insert(canon(name));
                        });
                    }
                    if let ExprKind::RefOf(place)=&expression.kind {
                        if let PlaceExpr::Ident(name)=place.as_ref() { excluded.insert(canon(name)); }
                    }
                }
                }
            }) { return Self::default(); }
        }
        bindings.excluded.extend(excluded.iter().cloned());
        for name in &excluded { bindings.claimed.insert(name.clone(), None); }
        bindings.by_binding.retain(|binding, _| !bindings.excluded.contains(binding));
        let eligible: HashSet<usize> = bindings.by_binding.values().copied().collect();
        bindings.declarations.retain(|ordinal, _| eligible.contains(ordinal));
        bindings
    }

    fn claim(&mut self, binding: String, owner: Option<usize>) {
        match self.claimed.entry(binding.clone()) {
            std::collections::hash_map::Entry::Vacant(entry) => { entry.insert(owner); }
            std::collections::hash_map::Entry::Occupied(mut entry) => {
                if owner.is_none() || *entry.get()!=owner {
                    entry.insert(None);
                    self.excluded.insert(binding.clone());
                    self.by_binding.remove(&binding);
                }
            }
        }
        if owner.is_none() { self.excluded.insert(binding.clone());self.by_binding.remove(&binding); }
    }

    fn begin(&mut self, ordinal: usize, statement: &Statement) {
        self.active = self.declarations.get(&ordinal).and_then(|declaration| {
            (declaration.statement == statement as *const Statement as usize).then_some(ordinal)
        });
    }
}

impl Compiler {
    /// Existing publisher policy, shared with the inventory. Prefixes encode
    /// publications here; consumers never decode names into producer identity.
    pub(super) fn source_function_publication_aliases(&self, name: &str) -> Vec<String> {
        if !self.profile.source_function_callable_aliases { return Vec::new(); }
        let canonical = format!("__vybe_func${}", self.canon(name));
        let mut aliases=vec![canonical.clone()];
        // Preserve the existing PHP publisher's case-folded compatibility alias.
        // This is not an admission policy or an optimization profile shortcut.
        if self.profile.name=="php" {
            let folded=format!("__vybe_func${}",
                crate::primitives::namespaces::normalize_source_path(name).to_ascii_lowercase());
            if folded!=canonical { aliases.push(folded); }
        }
        aliases.push(format!("__php_fn_{}",name.replace('.',"__")));
        aliases
    }

    pub(super) fn prepare_immutable_function_bindings(&mut self, body: &[Statement]) {
        self.immutable_function_bindings = if self.directives().declared_function_bindings_live == Some(false) {
            Bindings::prepare_with_aliases(body, |name| self.canon(name),
                |name| self.source_function_publication_aliases(name))
        } else { Bindings::default() };
    }

    pub(super) fn begin_immutable_function_declaration(&mut self, ordinal: usize, statement: &Statement) {
        self.immutable_function_bindings.begin(ordinal, statement);
    }

    pub(super) fn end_immutable_function_declaration(&mut self) {
        self.immutable_function_bindings.active = None;
    }

    pub(super) fn publish_immutable_function_alias(&mut self, name: &str, alias: &str,
        params: &[Param], body: &[Statement], chunk: usize) {
        if self.immutable_function_bindings.claimed.is_empty() { return; }
        let source=self.canon(name);
        let alias=self.canon(alias);
        let registry=&mut self.immutable_function_bindings;
        let owner=registry.by_binding.get(&source).copied();
        let valid=owner.is_some_and(|id| self.current==0 && registry.active==Some(id)
            && registry.claimed.get(&alias)==Some(&Some(id)) && !registry.excluded.contains(&alias)
            && registry.declarations.get(&id).is_some_and(|declaration|
                declaration.params==params.as_ptr() as usize && declaration.params_len==params.len()
                && declaration.body==body.as_ptr() as usize && declaration.body_len==body.len()
                && declaration.completed.is_some_and(|entry|entry.chunk==chunk)));
        if valid { registry.by_binding.insert(alias, owner.unwrap()); }
        else { registry.claim(alias, None); }
    }

    pub(super) fn complete_immutable_function_entry(
        &mut self, name: &str, params: &[Param], body: &[Statement], chunk: usize,
        adapted_entry: bool,
    ) {
        if adapted_entry || self.current != 0 { return; }
        let binding = self.canon(name);
        let registry = &mut self.immutable_function_bindings;
        let Some(id) = registry.by_binding.get(&binding).copied() else { return; };
        if registry.active != Some(id) { return; }
        let Some(declaration) = registry.declarations.get_mut(&id) else { return; };
        if declaration.params != params.as_ptr() as usize
            || declaration.body != body.as_ptr() as usize
            || declaration.params_len != params.len() || declaration.body_len != body.len()
        { return; }
        let callee = &self.chunks[chunk];
        if callee.direct_entry_has_captures != Some(false) || callee.is_async || callee.is_generator
            || callee.func_sig.is_some() || callee.native_func_type.is_some()
            || callee.result_arity != 1
            || usize::from(callee.arity) != params.len() + usize::from(callee.takes_receiver)
        { return; }
        declaration.completed = Some(CompletedEntry {
            chunk, arity: callee.arity, takes_receiver: callee.takes_receiver,
            results: callee.result_arity,
        });
    }

    pub(super) fn try_emit_immutable_declared_call(&mut self, name: &str, args: &[Argument]) -> Result<bool, String> {
        if self.directives().declared_function_bindings_live != Some(false)
            || self.current_class.is_some() || self.uses_proxy
            || self.has_accessible_local_binding(name)
            || args.iter().any(|arg| arg.by_ref || arg.spread || arg.name.is_some())
        { return Ok(false); }
        let binding = self.canon_lookup(name);
        let registry = &self.immutable_function_bindings;
        let entry = registry.by_binding.get(binding.as_ref())
            .and_then(|id| registry.declarations.get(id)).and_then(|decl| decl.completed);
        let Some(entry) = entry else { return Ok(false); };
        if entry.takes_receiver != self.universal_receiver()
            || usize::from(entry.arity) != args.len() + usize::from(entry.takes_receiver)
            || u32::try_from(entry.chunk).is_err()
        { return Ok(false); }
        let callee = &self.chunks[entry.chunk];
        if callee.direct_entry_has_captures != Some(false)
            || callee.arity != entry.arity || callee.result_arity != entry.results
            || callee.takes_receiver != entry.takes_receiver || callee.func_sig.is_some()
            || callee.native_func_type.is_some()
            || callee.is_async || callee.is_generator
        { return Ok(false); }
        // A thin copy of the actual completed header, not a guessed signature
        // and not a clone of the function's bytecode/constant pool per call.
        let mut signature = Chunk::new("<completed-source-function-entry>");
        signature.arity = entry.arity;
        signature.result_arity = entry.results;
        signature.direct_entry_has_captures = Some(false);
        let line = self.line;
        if entry.takes_receiver {
            crate::primitives::expressions::emit_undefined(self.chunk(), line);
        }
        for argument in args { self.compile_expr_with_value_copy(&argument.value)?; }
        self.chunk().emit_direct_defined_call(entry.chunk as u32, &signature, false, line)?;
        Ok(true)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn parameter(name: &str, ty: &str) -> Param {
        Param { name: name.into(), type_hint: Some(ty.into()), default: None,
            pass_by: PassBy::Value, is_rest: false, is_kwargs: false,
            is_optional: false, is_nullable: false }
    }
    fn function(name: &str, params: Vec<Param>, body: Vec<Statement>) -> Statement {
        Statement::new(StmtKind::FunctionDecl { name: name.into(), params,
            return_type: None, body, modifiers: Modifiers::default(), handles: Vec::new(),
            is_async: false, is_generator: false, is_sub: false })
    }
    fn call(name: &str) -> Expression {
        Expression::new(ExprKind::Call { callee: Box::new(Expression::ident(name)),
            args: vec![Argument::positional(Expression::int(3))], optional: false })
    }
    fn target() -> Statement {
        function("target", vec![parameter("value", "int32")],
            vec![Statement::new(StmtKind::Return(Some(Expression::ident("value"))))])
    }

    #[test]
    fn root_return_keeps_completed_declaration_candidates() {
        let body = vec![target(), Statement::new(StmtKind::Return(Some(call("target"))))];
        let bindings = Bindings::prepare(&body, str::to_owned);
        assert_eq!(bindings.by_binding.get("target"), Some(&0));
        assert_eq!(bindings.declarations.len(), 1);
    }

    #[test]
    fn named_function_expression_keeps_actual_completed_edges_across_publishers() {
        for text in [include_str!("../../../../languages/go/src/profile"),
            include_str!("../../../../languages/php/src/profile")] {
            let profile = crate::profile::parse_profile(text).unwrap();
            let guest = Expression::new(ExprKind::FunctionExpr(Box::new(
                function("guest", vec![], vec![Statement::new(StmtKind::Return(None))]))));
            let source = module(Some(false), vec![target(), caller(""),
                function("startup", vec![], vec![Statement::new(StmtKind::Expr(guest))])]);
            let chunks = Compiler::with_profile(profile).with_run_module_end_hooks(false)
                .compile(&source).unwrap();
            let target = chunks.iter().position(|chunk| chunk.name == "target").unwrap();
            let caller = chunks.iter().find(|chunk| chunk.name == "caller").unwrap();
            assert_eq!(caller.direct_defined_calls.values().next().map(|edge| edge.target as usize), Some(target));
        }
    }

    #[test]
    fn conditional_function_expression_alias_collision_blocks_actual_cross_profile_edge() {
        for text in [include_str!("../../../../languages/go/src/profile"),
            include_str!("../../../../languages/php/src/profile")] {
            let mut profile = crate::profile::parse_profile(text).unwrap();
            profile.source_function_callable_aliases = true;
            let conditional = Statement::new(StmtKind::If { cond: Expression::bool(false),
                then_body: vec![function("Outer.Inner", vec![], vec![])], elifs: vec![], else_body: None });
            let guest = Expression::new(ExprKind::FunctionExpr(Box::new(
                function("guest", vec![], vec![conditional]))));
            let source = module(Some(false), vec![
                function("__php_fn_Outer__Inner", vec![parameter("value", "int32")],
                    vec![Statement::new(StmtKind::Return(Some(Expression::ident("value"))))]),
                function("caller", vec![], vec![Statement::new(StmtKind::Return(Some(call("__php_fn_Outer__Inner"))))]),
                function("startup", vec![], vec![Statement::new(StmtKind::Expr(guest))]),
            ]);
            let chunks = Compiler::with_profile(profile).with_run_module_end_hooks(false)
                .compile(&source).unwrap();
            assert!(chunks.iter().find(|chunk| chunk.name == "caller").unwrap().direct_defined_calls.is_empty());
        }
    }

    fn module(policy: Option<bool>, body: Vec<Statement>) -> Module {
        Module { name: "immutable_function_test".into(), language: vybe_ast::Lang::Go,
            imports: Vec::new(), canon: Default::default(), body,
            directives: Directives { declared_function_bindings_live: policy,
                callable_invocation: Some(vybe_ast::CallableInvocation::Single),
                receiver_binding: Some(vybe_ast::ReceiverBinding::UniversalParameter),
                ..Default::default() } }
    }
    fn compile(module: &Module) -> Vec<Chunk> {
        let profile = crate::profile::parse_profile(
            include_str!("../../../../languages/go/src/profile")).unwrap();
        Compiler::with_profile(profile).with_run_module_end_hooks(false).compile(module).unwrap()
    }
    fn caller(param: &str) -> Statement {
        function("caller", if param.is_empty() { Vec::new() } else {
            vec![parameter(param, "func(int32) int32")] },
            vec![Statement::new(StmtKind::Return(Some(call("target"))))])
    }

    #[test]
    fn immutable_completed_declaration_emits_standard_direct_edge() {
        let chunks = compile(&module(Some(false), vec![target(), caller("")]));
        let target = chunks.iter().position(|chunk| chunk.name == "target").unwrap();
        let caller = chunks.iter().find(|chunk| chunk.name == "caller").unwrap();
        assert_eq!(caller.direct_defined_calls.len(), 1);
        let edge = caller.direct_defined_calls.values().next().unwrap();
        assert_eq!(edge.target as usize, target);
        assert!(edge.source);
        assert_eq!(chunks[target].direct_entry_has_captures, Some(false));
    }

    #[test]
    fn unstated_and_explicit_live_preserve_same_emission() {
        let a = compile(&module(None, vec![target(), caller("")]));
        let b = compile(&module(Some(true), vec![target(), caller("")]));
        assert_eq!(a.len(), b.len());
        for (a, b) in a.iter().zip(&b) {
            assert_eq!(a.code, b.code);
            assert_eq!(a.local_count, b.local_count);
            assert_eq!(a.direct_defined_calls, b.direct_defined_calls);
        }
    }

    #[test]
    fn lexical_function_parameter_shadow_does_not_use_declaration_edge() {
        let chunks = compile(&module(Some(false), vec![target(), caller("target")]));
        assert!(chunks.iter().find(|chunk| chunk.name == "caller").unwrap().direct_defined_calls.is_empty());
    }

    #[test]
    fn incomplete_recursive_and_forward_entries_keep_callable_path() {
        let recursive = function("target", vec![parameter("value", "int32")],
            vec![Statement::new(StmtKind::Return(Some(call("target"))))]);
        let chunks = compile(&module(Some(false), vec![recursive]));
        assert!(chunks.iter().find(|chunk| chunk.name == "target").unwrap().direct_defined_calls.is_empty());
        let chunks = compile(&module(Some(false), vec![caller(""), target()]));
        assert!(chunks.iter().find(|chunk| chunk.name == "caller").unwrap().direct_defined_calls.is_empty());
    }

    #[test]
    fn conditional_and_duplicate_declarations_are_not_certified() {
        let statements = vec![Statement::new(StmtKind::If {
            cond: Expression::int(1), then_body: vec![target()], elifs: Vec::new(), else_body: None,
        })];
        assert!(Bindings::prepare(&statements, str::to_owned).by_binding.is_empty());
        assert!(Bindings::prepare(&[target(), target()], str::to_owned).by_binding.is_empty());
    }

    #[test]
    fn source_statement_identity_not_equal_spelling_controls_activation() {
        let statements = vec![target()];
        let mut bindings = Bindings::prepare(&statements, str::to_owned);
        bindings.begin(0, &statements[0].clone());
        assert_eq!(bindings.active, None);
        bindings.begin(0, &statements[0]);
        assert_eq!(bindings.active, Some(0));
    }

    #[test]
    fn unknown_callback_has_no_completed_declaration_receipt() {
        let callback = || function("caller", vec![parameter("callback", "func(int32) int32")],
            vec![Statement::new(StmtKind::Return(Some(call("callback"))))]);
        let candidate = compile(&module(Some(false), vec![target(), callback()]));
        let legacy = compile(&module(None, vec![target(), callback()]));
        let candidate = candidate.iter().find(|chunk| chunk.name == "caller").unwrap();
        let legacy = legacy.iter().find(|chunk| chunk.name == "caller").unwrap();
        assert!(candidate.direct_defined_calls.is_empty());
        assert_eq!(candidate.code, legacy.code);
        assert_eq!(candidate.local_count, legacy.local_count);
    }

    #[test]
    fn only_completed_uncaptured_unadapted_native_producer_can_publish_entry() {
        let statements = vec![target()];
        let profile = crate::profile::parse_profile(
            include_str!("../../../../languages/go/src/profile")).unwrap();
        let mut compiler = Compiler::with_profile(profile);
        compiler.immutable_function_bindings = Bindings::prepare(&statements, str::to_owned);
        compiler.begin_immutable_function_declaration(0, &statements[0]);
        let StmtKind::FunctionDecl { params, body, .. } = &statements[0].kind else { unreachable!() };
        let mut chunk = Chunk::new("unrelated-debug-name");
        chunk.arity = 2; chunk.result_arity = 1; chunk.takes_receiver = true;
        let index = compiler.chunks.len();
        compiler.chunks.push(chunk);
        for capture_fact in [None, Some(true)] {
            compiler.chunks[index].direct_entry_has_captures = capture_fact;
            compiler.complete_immutable_function_entry("target", params, body, index, false);
            assert!(compiler.immutable_function_bindings.declarations[&0].completed.is_none());
        }
        compiler.chunks[index].direct_entry_has_captures = Some(false);
        compiler.complete_immutable_function_entry("target", params, body, index, true);
        assert!(compiler.immutable_function_bindings.declarations[&0].completed.is_none());
        compiler.complete_immutable_function_entry("target", params, body, index, false);
        assert_eq!(compiler.immutable_function_bindings.declarations[&0].completed.unwrap().chunk, index);
    }

    #[test]
    fn cloned_or_foreign_body_does_not_inherit_a_same_named_declaration_receipt() {
        let statements = vec![target()];
        let profile = crate::profile::parse_profile(
            include_str!("../../../../languages/go/src/profile")).unwrap();
        let mut compiler = Compiler::with_profile(profile);
        compiler.immutable_function_bindings = Bindings::prepare(&statements, str::to_owned);
        compiler.begin_immutable_function_declaration(0, &statements[0]);
        let StmtKind::FunctionDecl { params, body, .. } = &statements[0].kind else { unreachable!() };
        let copied_body = body.clone();
        let mut chunk = Chunk::new("target");
        chunk.arity = 2; chunk.result_arity = 1; chunk.takes_receiver = true;
        chunk.direct_entry_has_captures = Some(false);
        let index = compiler.chunks.len(); compiler.chunks.push(chunk);
        compiler.complete_immutable_function_entry("target", params, &copied_body, index, false);
        assert!(compiler.immutable_function_bindings.declarations[&0].completed.is_none());
    }

    #[test]
    fn known_root_control_retains_only_unconditional_candidates_and_writes_tombstone() {
        let conditional=Statement::new(StmtKind::If { cond:Expression::bool(false),
            then_body:vec![function("conditional",vec![],vec![])],elifs:vec![],else_body:None });
        let statements=vec![target(),conditional,Statement::new(StmtKind::Return(None))];
        let bindings=Bindings::prepare(&statements,str::to_owned);
        assert_eq!(bindings.by_binding.get("target"),Some(&0));
        assert!(!bindings.by_binding.contains_key("conditional"));
        let mut statements=statements;
        statements.push(Statement::new(StmtKind::Assign { targets:vec![Expression::ident("target")],
            value:Expression::int(0),by_ref:false }));
        let bindings=Bindings::prepare(&statements,str::to_owned);
        assert!(!bindings.by_binding.contains_key("target"));
        assert!(bindings.excluded.contains("target"));
    }

    #[test]
    fn planned_alias_collisions_and_conditional_producers_are_permanent_tombstones() {
        let statements=vec![target(),function("other",vec![],vec![])];
        let mut bindings=Bindings::prepare_with_aliases(&statements,str::to_owned,|_|vec!["shared".into()]);
        assert_eq!(bindings.claimed.get("shared"),Some(&None));
        bindings.claim("shared".into(),Some(0));
        assert_eq!(bindings.claimed.get("shared"),Some(&None));
        assert!(bindings.excluded.contains("shared"));
        let statements=vec![target(),Statement::new(StmtKind::If {cond:Expression::bool(false),
            then_body:vec![function("other",vec![],vec![])],elifs:vec![],else_body:None})];
        let bindings=Bindings::prepare_with_aliases(&statements,str::to_owned,|_|vec!["shared".into()]);
        assert!(bindings.excluded.contains("shared"));
    }

    #[test]
    fn actual_alias_publication_uses_completed_target_and_preserves_argument_copy() {
        let alias="__vybe_func$target";
        let statements=vec![target(),function("caller",vec![],
            vec![Statement::new(StmtKind::Return(Some(call(alias))))])];
        let profile=crate::profile::parse_profile(include_str!("../../../../languages/go/src/profile")).unwrap();
        let mut profile=profile;profile.source_function_callable_aliases=true;
        let chunks=Compiler::with_profile(profile).with_run_module_end_hooks(false)
            .compile(&module(Some(false),statements)).unwrap();
        let target=chunks.iter().position(|chunk|chunk.name=="target").unwrap();
        let caller=chunks.iter().find(|chunk|chunk.name=="caller").unwrap();
        assert_eq!(caller.direct_defined_calls.values().next().map(|edge|edge.target as usize),Some(target));
        // The alias path shares the exact argument-copy consumer, not a second
        // invocation emitter. Ordinary source and alias calls have equal IR.
        let mut profile=crate::profile::parse_profile(include_str!("../../../../languages/go/src/profile")).unwrap();
        profile.source_function_callable_aliases=true;
        let ordinary=Compiler::with_profile(profile).with_run_module_end_hooks(false)
            .compile(&module(Some(false),vec![super::tests::target(),super::tests::caller("")])).unwrap();
        let ordinary=ordinary.iter().find(|chunk|chunk.name=="caller").unwrap();
        assert_eq!(caller.code,ordinary.code);
    }

    #[test]
    fn alias_receipt_rejects_wrong_source_identity_and_local_shadow() {
        let alias="__vybe_func$target";
        let mut profile=crate::profile::parse_profile(include_str!("../../../../languages/go/src/profile")).unwrap();
        profile.source_function_callable_aliases=true;
        let statements=vec![target()];
        let mut compiler=Compiler::with_profile(profile.clone());
        compiler.prepare_immutable_function_bindings(&statements);
        // Explicitly prepare without depending on the current directive frame.
        compiler.immutable_function_bindings=Bindings::prepare_with_aliases(&statements,str::to_owned,
            |_|vec![alias.into()]);
        compiler.begin_immutable_function_declaration(0,&statements[0]);
        let StmtKind::FunctionDecl {params,body,..}=&statements[0].kind else{unreachable!()};
        let mut chunk=Chunk::new("diagnostic-only");chunk.arity=2;chunk.takes_receiver=true;
        chunk.result_arity=1;chunk.direct_entry_has_captures=Some(false);
        let index=compiler.chunks.len();compiler.chunks.push(chunk);
        compiler.complete_immutable_function_entry("target",params,body,index,false);
        let copy=body.clone();
        compiler.publish_immutable_function_alias("target",alias,params,&copy,index);
        assert!(!compiler.immutable_function_bindings.by_binding.contains_key(alias));
        compiler.publish_immutable_function_alias("target",alias,params,body,index);
        assert!(!compiler.immutable_function_bindings.by_binding.contains_key(alias));
        let caller=function("caller",vec![parameter(alias,"func(int32) int32")],
            vec![Statement::new(StmtKind::Return(Some(call(alias))))]);
        let chunks=Compiler::with_profile(profile).with_run_module_end_hooks(false)
            .compile(&module(Some(false),vec![target(),caller])).unwrap();
        assert!(chunks.iter().find(|chunk|chunk.name=="caller").unwrap().direct_defined_calls.is_empty());
    }

    #[test]
    fn alias_publication_without_a_completed_eligible_entry_never_admits() {
        for adapted in [false,true] {
            let statements=vec![target()];
            let profile=crate::profile::parse_profile(include_str!("../../../../languages/go/src/profile")).unwrap();
            let mut compiler=Compiler::with_profile(profile);
            compiler.immutable_function_bindings=Bindings::prepare_with_aliases(&statements,str::to_owned,
                |_|vec!["published".into()]);
            compiler.begin_immutable_function_declaration(0,&statements[0]);
            let StmtKind::FunctionDecl {params,body,..}=&statements[0].kind else{unreachable!()};
            let mut chunk=Chunk::new("target");chunk.arity=2;chunk.result_arity=1;chunk.takes_receiver=true;
            chunk.direct_entry_has_captures=Some(!adapted);
            let index=compiler.chunks.len();compiler.chunks.push(chunk);
            compiler.complete_immutable_function_entry("target",params,body,index,adapted);
            compiler.publish_immutable_function_alias("target","published",params,body,index);
            assert!(!compiler.immutable_function_bindings.by_binding.contains_key("published"));
            assert!(compiler.immutable_function_bindings.excluded.contains("published"));
        }
    }

    #[test]
    fn opaque_nested_publications_and_return_expression_writes_keep_fallback() {
        let statements=vec![target(),Statement::new(StmtKind::NamespaceDecl {
            name:"opaque".into(),body:vec![] })];
        assert!(Bindings::prepare(&statements,str::to_owned).by_binding.is_empty());
        let statements=vec![target(),Statement::new(StmtKind::Return(Some(Expression::new(ExprKind::Assign {
            target:Box::new(Expression::ident("target")),value:Box::new(Expression::int(0)),
        }))))];
        assert!(!Bindings::prepare(&statements,str::to_owned).by_binding.contains_key("target"));
    }

    #[test]
    fn published_alias_direct_edge_executes_through_existing_vm_call_transport() {
        let mut profile=crate::profile::parse_profile(include_str!("../../../../languages/go/src/profile")).unwrap();
        profile.source_function_callable_aliases=true;
        let chunks=Compiler::with_profile(profile).with_run_module_end_hooks(false).compile(
            &module(Some(false),vec![target(),Statement::new(StmtKind::Return(Some(call("__vybe_func$target"))))]))
            .unwrap();
        assert_eq!(chunks[0].direct_defined_calls.len(),1);
        let mut vm=vybe_runtime::VM::new();
        crate::primitives::platforms::register_platforms_all(&mut vm);
        crate::dynamic::register_dynamic_runtime_imports(&mut vm);
        match vm.run(chunks) {
            Ok(Value::I32(value)) => assert_eq!(value,3),
            Ok(Value::F64(value)) => assert_eq!(value,3.0),
            result => panic!("alias transport failed: {result:?}"),
        }
    }

    #[test]
    fn class_methods_are_not_global_producers_but_nested_publications_are_claimed() {
        let class=|body|Statement::new(StmtKind::ClassDecl {
            name:"Container".into(),parents:vec![],interfaces:vec![],
            members:vec![ClassMember::Method(Box::new(function("method",vec![],body)))],
            modifiers:Default::default(),decorators:vec![],
        });
        let statements=vec![target(),class(vec![Statement::new(StmtKind::Return(None))])];
        let bindings=Bindings::prepare_with_aliases(&statements,str::to_owned,|_|vec!["shared".into()]);
        assert_eq!(bindings.by_binding.get("target"),Some(&0));
        assert_eq!(bindings.claimed.get("shared"),Some(&Some(0)));
        assert!(!bindings.claimed.contains_key("method"));
        let statements=vec![target(),class(vec![function("nested",vec![],vec![])])];
        let bindings=Bindings::prepare_with_aliases(&statements,str::to_owned,|_|vec!["shared".into()]);
        assert!(bindings.excluded.contains("shared"));
        assert!(!bindings.by_binding.contains_key("nested"));
    }
}
