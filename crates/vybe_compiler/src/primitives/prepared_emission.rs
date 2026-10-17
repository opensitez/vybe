use super::*;

/// One immutable common-AST producer admitted by a prepared emission session.
/// Units can only be obtained from that session's owned producer iterator.
pub struct PreparedUnit<'a> {
    owner: Arc<()>,
    ordinal: usize,
    source_index: usize,
    statement: &'a mut Statement,
    retired: &'a mut bool,
}

impl PreparedUnit<'_> {
    pub fn source_index(&self) -> usize { self.source_index }
    pub fn statement(&self) -> &Statement { self.statement }
    pub fn is_function_declaration(&self) -> bool {
        matches!(self.statement.kind, StmtKind::FunctionDecl { .. })
    }
}

#[derive(Default)]
struct UnitCursor {
    function: usize,
    body: usize,
}

impl UnitCursor {
    fn next(&mut self, count: usize, mut is_function: impl FnMut(usize) -> Option<bool>) -> Option<usize> {
        while self.function < count {
            let index = self.function;
            self.function += 1;
            if is_function(index) == Some(true) { return Some(index); }
        }
        while self.body < count {
            let index = self.body;
            self.body += 1;
            if is_function(index) == Some(false) { return Some(index); }
        }
        None
    }
}

/// Lends prepared slots in function-first order without moving their ASTs.
/// Each unit must be consumed before another can be requested from this producer.
pub struct PreparedUnitProducer {
    owner: Arc<()>,
    body: Vec<Statement>,
    retired: Vec<bool>,
    builtin_signatures: Option<Statement>,
    builtin_retired: bool,
    cursor: UnitCursor,
    ordinal: usize,
    count: usize,
}

impl PreparedUnitProducer {
    fn new(body: Vec<Statement>, builtin_signatures: Option<Statement>) -> Self {
        let count = body.len() + usize::from(builtin_signatures.is_some());
        let retired = vec![false; body.len()];
        Self { owner: Arc::new(()), body, retired, builtin_signatures,
            builtin_retired: false, cursor: UnitCursor::default(), ordinal: 0, count }
    }

    pub fn next_unit(&mut self) -> Option<PreparedUnit<'_>> {
        let body = &self.body;
        let retired = &self.retired;
        let builtin = &self.builtin_signatures;
        let index = self.cursor.next(self.count, |index| {
            let stmt = if index < body.len() {
                if retired[index] { return None; }
                &body[index]
            } else {
                if self.builtin_retired { return None; }
                builtin.as_ref()?
            };
            Some(matches!(stmt.kind, StmtKind::FunctionDecl { .. }))
        })?;
        let (statement, retired) = if index < self.body.len() {
            (&mut self.body[index], &mut self.retired[index])
        } else {
            (self.builtin_signatures.as_mut().unwrap(), &mut self.builtin_retired)
        };
        let ordinal = self.ordinal;
        self.ordinal += 1;
        Some(PreparedUnit { owner: self.owner.clone(), ordinal, source_index: index, statement, retired })
    }

    pub fn remaining_units(&self) -> usize { self.count - self.ordinal }

    pub fn retained_statements(&self) -> usize {
        self.retired.iter().filter(|retired| !**retired).count()
            + usize::from(self.builtin_signatures.is_some() && !self.builtin_retired)
    }
}

struct BorrowedUnits<'a> {
    body: &'a [Statement],
    builtin_signatures: Option<Statement>,
    cursor: UnitCursor,
}

impl BorrowedUnits<'_> {
    fn next(&mut self) -> Option<(usize, &Statement)> {
        let body = self.body;
        let builtin = &self.builtin_signatures;
        let count = body.len() + usize::from(builtin.is_some());
        let index = self.cursor.next(count, |index| {
            let stmt = if index < body.len() { &body[index] } else { builtin.as_ref()? };
            Some(matches!(stmt.kind, StmtKind::FunctionDecl { .. }))
        })?;
        Some((index, if index < body.len() { &body[index] }
            else { self.builtin_signatures.as_ref().unwrap() }))
    }
}

enum PreparedBody<'a> {
    Borrowed(BorrowedUnits<'a>),
    Owned(PreparedUnitProducer),
    Detached,
}

enum NextUnit<'a> {
    Borrowed(usize, &'a Statement),
    Owned(PreparedUnit<'a>),
}

impl PreparedBody<'_> {
    fn next(&mut self) -> Option<NextUnit<'_>> {
        match self {
            Self::Borrowed(producer) => producer.next().map(|(index, stmt)| NextUnit::Borrowed(index, stmt)),
            Self::Owned(producer) => producer.next_unit().map(NextUnit::Owned),
            Self::Detached => None,
        }
    }

    fn function_count(&self) -> usize {
        match self {
            Self::Borrowed(producer) => producer.body.iter()
                .filter(|stmt| matches!(stmt.kind, StmtKind::FunctionDecl { .. })).count(),
            Self::Owned(producer) => producer.body.iter().zip(&producer.retired)
                .filter(|(stmt, retired)| !**retired && matches!(stmt.kind, StmtKind::FunctionDecl { .. })).count(),
            Self::Detached => 0,
        }
    }
}

pub(super) struct StatementStreamParts<'a> {
    pub compiler: Compiler,
    pub body: std::borrow::Cow<'a, [Statement]>,
    pub builtin_signatures: Option<Statement>,
    pub component_source: bool,
    pub component_binary: Option<Arc<vybe_ast::canon::ComponentBinary>>,
    pub timing: Option<String>,
    pub total_started: Instant,
}

struct EmissionState {
    strict_after_functions: bool,
    functions_finished: bool,
    body_finished: bool,
    function_pass_index: usize,
    function_elapsed: Duration,
    body_elapsed: Duration,
    module_producers: Option<std::collections::BTreeMap<&'static str, (usize, Duration)>>,
    timing: Option<String>,
}

impl EmissionState {
    fn finish_functions(&mut self, compiler: &mut Compiler) {
        if self.functions_finished { return; }
        self.functions_finished = true;
        if let Some(path) = self.timing.as_deref() {
            write_compiler_timing(path, "compile_function_decls", self.function_elapsed);
        }
        if self.strict_after_functions { compiler.in_strict = true; }
    }

    fn emit(&mut self, compiler: &mut Compiler, index: usize, stmt: &Statement) -> Result<EmissionProgress, String> {
        let _debug_compile = vybe_runtime::debugger::DebugPhase::current("compiler total");
        let _source_lines = vybe_runtime::source_lines::LineEmissionScope::enter(compiler.compact_source_lines);
        let chunks = compiler.chunks.len();
        let bytes = compiler.chunks[0].code.len();
        let function = matches!(stmt.kind, StmtKind::FunctionDecl { .. });
        if let StmtKind::FunctionDecl { name, .. } = &stmt.kind {
            let _debug_functions = vybe_runtime::debugger::DebugPhase::current("compiler functions");
            let _debug_function = vybe_runtime::debugger::DebugPhase::current_lazy(|| format!("function {name}"));
            let started = Instant::now();
            if let Some(path) = self.timing.as_deref() {
                write_compiler_timing(path, &format!("compile_function_start[{}] {name}", self.function_pass_index), Duration::ZERO);
            }
            compiler.begin_immutable_function_declaration(index, stmt);
            let emitted = compiler.compile_stmt(stmt);
            compiler.end_immutable_function_declaration();
            emitted?;
            let elapsed = started.elapsed();
            self.function_elapsed += elapsed;
            if let Some(path) = self.timing.as_deref() {
                self.function_pass_index += 1;
                write_compiler_timing(path, &format!("compile_function_decl[{}] {name}", self.function_pass_index), elapsed);
            }
        } else {
            self.finish_functions(compiler);
            let _debug_body = vybe_runtime::debugger::DebugPhase::current("compiler module body");
            let started = self.timing.as_ref().map(|_| Instant::now());
            compiler.compile_stmt(stmt)?;
            if let (Some(path), Some(started)) = (self.timing.as_deref(), started) {
                let elapsed = started.elapsed();
                self.body_elapsed += elapsed;
                let kind = match &stmt.kind {
                    StmtKind::ClassDecl { .. } => "class", StmtKind::Expr(_) => "expression",
                    StmtKind::VarDecl { .. } => "variable", StmtKind::If { .. } => "conditional",
                    StmtKind::Block(_) => "block", StmtKind::Try { .. } => "try", _ => "other",
                };
                let aggregate = self.module_producers.as_mut().unwrap().entry(kind).or_default();
                aggregate.0 += 1; aggregate.1 += elapsed;
                if let StmtKind::ClassDecl { name, .. } = &stmt.kind {
                    write_compiler_timing(path, &format!("compile_class_decl {name} line={}", stmt.span.start_line), elapsed);
                }
            }
        }
        Ok(EmissionProgress { source_index: index, function_declaration: function,
            appended_chunks: chunks..compiler.chunks.len(), entry_bytes: bytes..compiler.chunks[0].code.len() })
    }

    fn complete(&mut self, compiler: &mut Compiler) {
        self.finish_functions(compiler);
        if self.body_finished { return; }
        self.body_finished = true;
        if let Some(path) = self.timing.as_deref() {
            for (kind, (count, elapsed)) in self.module_producers.take().unwrap() {
                write_compiler_timing(path, &format!("compile_module_producer_kind {kind} count={count}"), elapsed);
            }
            write_compiler_timing(path, "compile_non_function_stmts", self.body_elapsed);
        }
    }
}

/// One validated source unit, with all global preparation complete.
/// Owned producers are released after emission; borrowed source ASTs stay intact.
/// Pending chunks remain unpublished until the unit's final relocation succeeds.
pub struct PreparedEmission<'a> {
    compiler: Option<Compiler>,
    body: PreparedBody<'a>,
    owner: Option<Arc<()>>,
    count: usize,
    emitted: usize,
    state: EmissionState,
    component_source: bool,
    component_binary: Option<Arc<vybe_ast::canon::ComponentBinary>>,
    total_started: Instant,
}

impl<'a> PreparedEmission<'a> {
    pub(super) fn new(mut compiler: Compiler, body: std::borrow::Cow<'a, [Statement]>,
        builtin_signatures: Option<Statement>, component_source: bool,
        component_binary: Option<Arc<vybe_ast::canon::ComponentBinary>>,
        timing: Option<String>, total_started: Instant) -> Self {
        let strict_after_functions = compiler.profile.ecma_strict_mode
            && Compiler::stmts_have_use_strict_directive(&body);
        compiler.prepare_immutable_function_bindings(body.as_ref());
        let count = body.len() + usize::from(builtin_signatures.is_some());
        let body = match body {
            std::borrow::Cow::Borrowed(body) => PreparedBody::Borrowed(BorrowedUnits {
                body, builtin_signatures, cursor: UnitCursor::default(),
            }),
            std::borrow::Cow::Owned(body) => PreparedBody::Owned(PreparedUnitProducer::new(body, builtin_signatures)),
        };
        let owner = match &body { PreparedBody::Owned(producer) => Some(producer.owner.clone()), _ => None };
        let state = EmissionState { strict_after_functions, functions_finished: false, body_finished: false,
            function_pass_index: 0, function_elapsed: Duration::ZERO, body_elapsed: Duration::ZERO,
            module_producers: timing.as_ref().map(|_| std::collections::BTreeMap::new()), timing };
        Self { compiler: Some(compiler), body, owner, count, emitted: 0, state,
            component_source, component_binary, total_started }
    }

    pub fn pending_chunks(&self) -> Result<&[Chunk], String> {
        self.compiler.as_ref().map(|compiler| compiler.chunks.as_slice())
            .ok_or_else(|| "emission session is poisoned".to_string())
    }

    pub(super) fn into_statement_stream_parts(mut self) -> Result<StatementStreamParts<'a>, String> {
        if self.emitted != 0 || matches!(self.body, PreparedBody::Detached) {
            return Err("statement stream handoff requires an unstarted emission session".to_string());
        }
        for _ in 0..self.body.function_count() {
            let progress = self.emit_next()?.ok_or("missing hoisted function producer")?;
            if !progress.function_declaration { return Err("invalid hoisted function order".to_string()); }
        }
        let mut compiler = self.compiler.take().ok_or("emission session is poisoned")?;
        self.state.finish_functions(&mut compiler);
        compiler.source_atomic_free_bodies.clear();
        let (body, builtin_signatures) = match self.body {
            PreparedBody::Borrowed(producer) =>
                (std::borrow::Cow::Borrowed(producer.body), producer.builtin_signatures),
            PreparedBody::Owned(producer) =>
                (std::borrow::Cow::Owned(producer.body), producer.builtin_signatures),
            PreparedBody::Detached => unreachable!(),
        };
        Ok(StatementStreamParts { compiler, body, builtin_signatures,
            component_source: self.component_source, component_binary: self.component_binary,
            timing: self.state.timing, total_started: self.total_started })
    }

    /// Transfer only already-prepared ASTs. Units are session-bound and must be
    /// emitted in iterator order; they cannot be mutated or replaced by callers.
    pub fn take_owned_units(&mut self) -> Result<PreparedUnitProducer, String> {
        let compiler = self.compiler.as_mut().ok_or_else(|| "emission session is poisoned".to_string())?;
        match std::mem::replace(&mut self.body, PreparedBody::Detached) {
            PreparedBody::Owned(producer) => {
                // The caller may drop any remaining unit before emitting another.
                // Address identities cannot outlive that independent ownership.
                compiler.source_atomic_free_bodies.clear();
                Ok(producer)
            }
            other => {
                self.body = other;
                Err("emission session has no attached owned producer".to_string())
            }
        }
    }

    /// Consume a prepared unit from this session's detached producer.
    pub fn emit_unit(&mut self, unit: PreparedUnit<'_>) -> Result<EmissionProgress, String> {
        let result = if !matches!(self.body, PreparedBody::Detached) {
            Err("owned unit producer is not detached".to_string())
        } else if let Some(compiler) = self.compiler.as_mut() {
            emit_owned(compiler, &mut self.state, self.owner.as_ref(), &mut self.emitted, unit)
        } else { Err("emission session is poisoned".to_string()) };
        if result.is_err() { self.compiler.take(); }
        result
    }

    pub fn emit_next(&mut self) -> Result<Option<EmissionProgress>, String> {
        let result = self.emit_next_inner();
        if result.is_err() { self.compiler.take(); }
        result
    }

    fn emit_next_inner(&mut self) -> Result<Option<EmissionProgress>, String> {
        let compiler = self.compiler.as_mut().ok_or_else(|| "emission session is poisoned".to_string())?;
        match self.body.next() {
            Some(NextUnit::Borrowed(index, statement)) => {
                let progress = self.state.emit(compiler, index, statement)?;
                self.emitted += 1;
                Ok(Some(progress))
            }
            Some(NextUnit::Owned(unit)) =>
                emit_owned(compiler, &mut self.state, self.owner.as_ref(), &mut self.emitted, unit).map(Some),
            None if self.emitted == self.count => {
                self.state.complete(compiler);
                Ok(None)
            }
            None => Err("owned unit producer has pending statements".to_string()),
        }
    }

    pub fn finish(mut self) -> Result<CompileResult, String> {
        while self.emit_next()?.is_some() {}
        self.compiler.take().ok_or_else(|| "emission session is poisoned".to_string())?
            .finish_emission(self.component_source, self.component_binary, self.state.timing, self.total_started)
    }

    pub fn abort(self) {}
}

fn emit_owned(compiler: &mut Compiler, state: &mut EmissionState, owner: Option<&Arc<()>>,
    emitted: &mut usize, unit: PreparedUnit<'_>) -> Result<EmissionProgress, String> {
    if !owner.is_some_and(|owner| Arc::ptr_eq(owner, &unit.owner)) || unit.ordinal != *emitted {
        return Err("prepared unit does not match this session's next producer".to_string());
    }
    let result = state.emit(compiler, unit.source_index, unit.statement);
    if result.is_ok() {
        compiler.retire_source_body_receipts(unit.statement);
        *unit.statement = Statement::new(StmtKind::Empty);
        *unit.retired = true;
        *emitted += 1;
    }
    result
}
