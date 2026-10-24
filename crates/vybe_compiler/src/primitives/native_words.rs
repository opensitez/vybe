//! Native word locals: a binding the native-storage proof admits lives in a
//! physical i64 Wasm local of its declared lane. Values stay raw through
//! exact W64 operands and cross the value ABI on the exact carrier only where
//! a read leaves the native lane. See `documentation/word_local_transport_contract.md`.
use super::*;
use crate::primitives::scope::NativeWord;
use vybe_ast::word_local_transport::WordInterpretation;

impl Compiler {
    /// The native lane a declaration resolves to: a 32- or 64-bit integer
    /// width the profile declares for that type spelling.
    pub(super) fn native_word_reading(
        &self,
        hint: &vybe_ast::TypeHint,
    ) -> Option<(vybe_ast::BitLane, WordInterpretation)> {
        let width = self
            .profile
            .integer_cast_widths
            .get(self.canon_lookup(hint).as_ref())?;
        let lane = match width.bits {
            Some(64) => vybe_ast::BitLane::W64,
            Some(32) => vybe_ast::BitLane::W32,
            _ => return None,
        };
        Some((lane, if width.signed { WordInterpretation::Signed } else { WordInterpretation::Unsigned }))
    }

    /// The 64-bit reading of a declaration: the lanes a typed entry passes raw.
    fn native_entry_reading(&self, hint: &vybe_ast::TypeHint) -> Option<WordInterpretation> {
        self.native_word_reading(hint)
            .filter(|(lane, _)| *lane == vybe_ast::BitLane::W64)
            .map(|(_, reading)| reading)
    }

    /// Record which names of the function being entered keep native storage.
    /// Address-taken, atomic and closure-captured bindings stay boxed.
    pub(super) fn admit_native_words(&mut self, params: &[Param], body: &[Statement]) {
        let proof = crate::word_local_proof::prove_native_subset(
            self.directives(),
            params,
            body,
            |hint| self.native_word_reading(hint),
        );
        let admitted: Vec<_> = proof
            .eligible
            .into_iter()
            .filter(|binding| {
                !self.current_addr_taken_locals.contains(&binding.name)
                    && !self.current_atomic_word_locals.contains(&binding.name)
                    && !self.current_closure_captured_locals.contains(&binding.name)
            })
            .map(|binding| {
                (binding.name, NativeWord { lane: binding.lane, reading: binding.interpretation })
            })
            .collect();
        self.scope_mut().native_word_bindings.extend(admitted);
    }

    /// A fresh slot typed i64 for the whole function. The slot is above every
    /// scratch slot used so far, and declared slots are never reclaimed, so no
    /// externref ever shares it.
    fn define_native_word_local(
        &mut self,
        name: &str,
        hint: Option<vybe_ast::TypeHint>,
        word: NativeWord,
    ) -> u16 {
        let chunk = &mut self.chunks[self.current];
        chunk.local_count = chunk.local_count.max(chunk.scratch_high_water);
        let slot = self.define_source_local_typed(name, hint);
        self.chunks[self.current].i64_locals.insert(slot);
        self.scope_mut().native_word_slots.insert(slot, word);
        slot
    }

    /// `value` as the raw i64 a native slot of `word` holds: a W32 word is
    /// wrapped to its lane and extended by its reading.
    fn compile_native_word_value(&mut self, word: NativeWord, value: &Expression) -> Result<(), String> {
        if word.lane == vybe_ast::BitLane::W64 {
            return self.compile_exact_i64_operand(value);
        }
        let value = self.w32_cast_operand(value);
        if let ExprKind::Lit(Literal::Int(constant)) = &value.kind {
            let wrapped = *constant as i32;
            let line = self.line;
            self.chunk().emit_i64_const(
                if word.signed() { i64::from(wrapped) } else { i64::from(wrapped as u32) },
                line,
            );
            return Ok(());
        }
        // A W32 integer operation leaves its i32 word as its value.
        let word_result = matches!(&value.kind,
            ExprKind::Binary { op: BinOp::Integer(operation, vybe_ast::BitLane::W32), .. }
                if !operation.is_comparison());
        if self.compile_char_widening(value)? || (!self.compile_word_operand(value)? && !word_result) {
            let line = self.line;
            crate::primitives::bits::emit_normalize_integer(self.chunk(), vybe_ast::BitLane::W32, line);
        }
        self.emit_native_word_extend(word);
        Ok(())
    }

    /// The operand of a 32-bit integer cast whose bits are its operand's own
    /// low 32 bits: an integer constant or a W32 integer operation. Any other
    /// expression is returned as is.
    fn w32_cast_operand<'e>(&self, value: &'e Expression) -> &'e Expression {
        let ExprKind::Cast { expr, type_name } = &value.kind else {
            return value;
        };
        let w32 = self
            .profile
            .integer_cast_widths
            .get(self.canon_lookup(type_name).as_ref())
            .is_some_and(|width| width.bits == Some(32));
        let integer = matches!(&expr.kind, ExprKind::Lit(Literal::Int(_)))
            || matches!(&expr.kind,
                ExprKind::Binary { op: BinOp::Integer(operation, vybe_ast::BitLane::W32), .. }
                    if !operation.is_comparison());
        if w32 && integer { expr } else { value }
    }

    /// Stack `[i32 word]` -> `[i64]` extended by `word`'s reading.
    fn emit_native_word_extend(&mut self, word: NativeWord) {
        self.emit(if word.signed() { Op::I64_EXTEND_I32_S } else { Op::I64_EXTEND_I32_U });
    }

    /// Stack `[value]` -> the raw i64 a native slot of `word` holds.
    fn emit_native_word_unbox(&mut self, word: NativeWord) {
        let line = self.line;
        if word.lane == vybe_ast::BitLane::W32 {
            crate::primitives::bits::emit_normalize_integer(self.chunk(), vybe_ast::BitLane::W32, line);
            self.emit_native_word_extend(word);
        } else {
            crate::primitives::exact_integer::emit_value_to_i64(
                &mut self.chunks,
                self.current,
                word.signed(),
                line,
            );
        }
    }

    /// Stack `[raw i64 of word]` -> its value. A W32 word is always an exact
    /// Number; a W64 word takes the region's exact carrier.
    fn emit_native_word_box(&mut self, word: NativeWord) {
        let line = self.line;
        if word.lane == vybe_ast::BitLane::W32 {
            self.emit(Op::F64_CONVERT_I64_S);
            return;
        }
        let transport = self.directives().wide_integer_transport();
        crate::primitives::exact_integer::emit_box_i64_as(
            &mut self.chunks,
            self.current,
            word.signed(),
            transport,
            line,
        );
    }

    /// `decl` of an admitted name: its initializer as a raw word, stored.
    pub(super) fn compile_native_word_declaration(
        &mut self,
        name: &str,
        decl: &VarDeclarator,
    ) -> Result<bool, String> {
        let Some(word) = self.scope().native_word_bindings.get(name).copied() else {
            return Ok(false);
        };
        match &decl.init {
            Some(init) => self.compile_native_word_value(word, init)?,
            None => {
                let line = self.line;
                self.chunk().emit_i64_const(0, line);
            }
        }
        let slot = self.define_native_word_local(name, decl.type_hint.clone(), word);
        self.emit_u16(Op::LOCAL_SET, slot);
        Ok(true)
    }

    /// Rebind each admitted by-value parameter to a native slot holding its
    /// argument, unboxed once at entry.
    pub(super) fn bind_native_word_params(&mut self, params: &[Param]) {
        for param in params {
            let Some(word) = self.scope().native_word_bindings.get(&param.name).copied() else {
                continue;
            };
            if self.scope().native_word_slot(&param.name).is_some() {
                continue;
            }
            let Some(argument) = self.scope().resolve(&param.name) else {
                continue;
            };
            self.emit_u16(Op::LOCAL_GET, argument);
            self.emit_native_word_unbox(word);
            let slot = self.define_native_word_local(&param.name, param.type_hint.clone(), word);
            self.emit_u16(Op::LOCAL_SET, slot);
        }
    }

    /// Read `name` if it is a native word: the raw word boxed at its
    /// declared lane and reading.
    pub(super) fn emit_native_word_get(&mut self, name: &str) -> bool {
        let Some((slot, word)) = self.scope().native_word_slot(name) else {
            return false;
        };
        self.emit_u16(Op::LOCAL_GET, slot);
        self.emit_native_word_box(word);
        true
    }

    /// Read `name` as an i32 word if it is a native W32 word.
    pub(super) fn emit_native_w32_word(&mut self, name: &str) -> bool {
        let Some((slot, word)) = self.scope().native_word_slot(name) else {
            return false;
        };
        if word.lane != vybe_ast::BitLane::W32 {
            return false;
        }
        self.emit_u16(Op::LOCAL_GET, slot);
        self.emit(Op::I32_WRAP_I64);
        true
    }

    /// Store the boxed value on the stack into `name` if it is a native word.
    pub(super) fn emit_native_word_set_from_value(&mut self, name: &str) -> bool {
        let Some((slot, word)) = self.scope().native_word_slot(name) else {
            return false;
        };
        self.emit_native_word_unbox(word);
        self.emit_u16(Op::LOCAL_SET, slot);
        true
    }

    /// `name++` / `name--` with the value discarded, if `name` is a native word.
    pub(super) fn emit_native_word_step(&mut self, target: &Expression, increment: bool) -> bool {
        let ExprKind::Ident(name) = &target.kind else {
            return false;
        };
        let Some((slot, word)) = self.scope().native_word_slot(name) else {
            return false;
        };
        let line = self.line;
        self.emit_u16(Op::LOCAL_GET, slot);
        self.chunk().emit_i64_const(1, line);
        self.emit(if increment { Op::I64_ADD } else { Op::I64_SUB });
        if word.lane == vybe_ast::BitLane::W32 {
            self.emit(Op::I32_WRAP_I64);
            self.emit_native_word_extend(word);
        }
        self.emit_u16(Op::LOCAL_SET, slot);
        true
    }

    /// `target = value` when `target` names a native word: the value compiled
    /// as a raw word and stored. With `observed`, the stored word is also left
    /// on the stack on the exact carrier.
    pub(super) fn compile_native_word_assignment(
        &mut self,
        target: &Expression,
        value: &Expression,
        observed: bool,
    ) -> Result<bool, String> {
        let ExprKind::Ident(name) = &target.kind else {
            return Ok(false);
        };
        let Some((slot, word)) = self.scope().native_word_slot(name) else {
            return Ok(false);
        };
        self.compile_native_word_value(word, value)?;
        if observed {
            self.emit_u16(Op::LOCAL_TEE, slot);
            self.emit_native_word_box(word);
        } else {
            self.emit_u16(Op::LOCAL_SET, slot);
        }
        Ok(true)
    }

    /// The raw word of a native binding `value`, if it is one.
    pub(super) fn emit_native_word_operand(&mut self, value: &Expression) -> bool {
        let ExprKind::Ident(name) = &value.kind else {
            return false;
        };
        let Some((slot, _)) = self.scope().native_word_slot(name) else {
            return false;
        };
        self.emit_u16(Op::LOCAL_GET, slot);
        true
    }
}

/// A source function compiled with a typed physical entry: native word
/// parameters and result cross as raw i64, the rest as externref.
#[derive(Debug, Clone)]
pub(crate) struct NativeEntry {
    pub chunk: usize,
    pub params: Vec<Option<WordInterpretation>>,
    pub result: Option<WordInterpretation>,
}

fn has_handler(body: &[Statement]) -> bool {
    body.iter().any(|statement| match &statement.kind {
        StmtKind::Try { .. } | StmtKind::Select { .. } => true,
        StmtKind::Block(body)
        | StmtKind::While { body, .. }
        | StmtKind::DoWhile { body, .. }
        | StmtKind::For { body, .. }
        | StmtKind::ForIn { body, .. } => has_handler(body),
        StmtKind::Labeled { body, .. } => has_handler(std::slice::from_ref(body)),
        StmtKind::If { then_body, elifs, else_body, .. } => {
            has_handler(then_body)
                || elifs.iter().any(|(_, body)| has_handler(body))
                || else_body.as_deref().is_some_and(has_handler)
        }
        StmtKind::Switch { cases, default, .. } => {
            cases.iter().any(|case| has_handler(&case.body))
                || default.as_deref().is_some_and(has_handler)
        }
        _ => false,
    })
}

/// A typed body is a non-source Wasm function to the writer, which gives a
/// closure value no source-callable carrier there; such bodies keep the value
/// ABI.
fn creates_closure(body: &[Statement]) -> bool {
    let mut found = false;
    for statement in body {
        statement.walk_exprs(&mut |value| {
            found |= matches!(value.kind, ExprKind::Lambda { .. } | ExprKind::FunctionExpr(_));
        });
    }
    found
}

impl Compiler {
    /// The typed entry shape of a module-level function, or None when every
    /// parameter and the result are values. Handler regions (try/select)
    /// keep the value ABI: their unwinding paths return values. So do bodies
    /// that create closures (see [`creates_closure`]).
    pub(super) fn native_entry_shape(
        &self,
        params: &[Param],
        return_type: &Option<String>,
        body: &[Statement],
    ) -> Option<(Vec<Option<WordInterpretation>>, Option<WordInterpretation>)> {
        if self.directives().word_local_transport()
            != vybe_ast::WordLocalTransport::ProvenNative
            || has_handler(body)
            || creates_closure(body)
        {
            return None;
        }
        let reading = |hint: Option<&vybe_ast::TypeHint>| {
            hint.filter(|hint| {
                hint.binding == vybe_ast::TypeBinding::Checked
                    && hint.access == Some(vybe_ast::StorageAccess::Ordinary)
            })
            .and_then(|hint| self.native_entry_reading(hint))
        };
        let shape: Vec<_> = params
            .iter()
            .map(|param| {
                (param.pass_by == PassBy::Value && !param.is_rest && param.default.is_none())
                    .then(|| reading(param.type_hint.as_ref()))
                    .flatten()
            })
            .collect();
        let result = return_type.as_ref().and_then(|spelling| {
            self.native_entry_reading(&vybe_ast::TypeHint::checked(spelling.trim()))
        });
        (result.is_some() || shape.iter().any(Option::is_some)).then_some((shape, result))
    }

    /// The Wasm signature of a typed entry, receiver first when it has one.
    pub(super) fn native_entry_signature(
        params: &[Option<WordInterpretation>],
        result: Option<WordInterpretation>,
        receiver: bool,
    ) -> (String, String) {
        let mut types: Vec<&str> = Vec::new();
        if receiver {
            types.push("externref");
        }
        types.extend(params.iter().map(|param| if param.is_some() { "i64" } else { "externref" }));
        (types.join(","), if result.is_some() { "i64" } else { "externref" }.to_owned())
    }

    /// At entry of a typed function: admitted native parameters keep their
    /// i64 argument slot; the others are boxed once into a value local.
    pub(super) fn bind_native_entry_params(
        &mut self,
        params: &[Param],
        shape: &[Option<WordInterpretation>],
    ) {
        for (param, reading) in params.iter().zip(shape) {
            let Some(reading) = *reading else { continue };
            let Some(slot) = self.scope().resolve(&param.name) else { continue };
            self.chunks[self.current].i64_locals.insert(slot);
            let word = NativeWord { lane: vybe_ast::BitLane::W64, reading };
            if self.scope().native_word_bindings.get(&param.name) == Some(&word) {
                self.scope_mut().native_word_slots.insert(slot, word);
            } else {
                let line = self.line;
                self.emit_u16(Op::LOCAL_GET, slot);
                let transport = self.directives().wide_integer_transport();
                crate::primitives::exact_integer::emit_box_i64_as(
                    &mut self.chunks,
                    self.current,
                    reading == WordInterpretation::Signed,
                    transport,
                    line,
                );
                let boxed = self.define_source_local_typed(&param.name, param.type_hint.clone());
                self.emit_u16(Op::LOCAL_SET, boxed);
            }
        }
    }

    /// The value-ABI entry of typed function `entry`: unbox native arguments,
    /// call the typed body directly, box a native result.
    pub(super) fn build_native_entry_wrapper(
        &mut self,
        name: &str,
        entry: &NativeEntry,
        arity: u8,
        receiver: bool,
    ) -> Result<usize, String> {
        let wrapper = self.chunks.len();
        let mut chunk = common::functions::create_function_chunk(name, arity);
        chunk.takes_receiver = self.chunks[entry.chunk].takes_receiver;
        chunk.result_arity = self.chunks[entry.chunk].result_arity;
        chunk.param_count = self.chunks[entry.chunk].param_count;
        chunk.local_count = u16::from(arity);
        chunk.scratch_high_water = u16::from(arity);
        self.chunks.push(chunk);
        let line = self.line;
        let offset = u16::from(receiver);
        if receiver {
            self.chunks[wrapper].emit_op_u16(Op::LOCAL_GET, 0, line);
        }
        for (index, reading) in entry.params.iter().enumerate() {
            self.chunks[wrapper].emit_op_u16(Op::LOCAL_GET, index as u16 + offset, line);
            if let Some(reading) = reading {
                let signed = *reading == WordInterpretation::Signed;
                crate::primitives::exact_integer::emit_value_to_i64(&mut self.chunks, wrapper, signed, line);
            }
        }
        common::functions::emit_direct_defined_call(&mut self.chunks, wrapper, entry.chunk, false, line)?;
        if let Some(reading) = entry.result {
            let transport = self.directives().wide_integer_transport();
            crate::primitives::exact_integer::emit_box_i64_as(
                &mut self.chunks,
                wrapper,
                reading == WordInterpretation::Signed,
                transport,
                line,
            );
        }
        self.chunks[wrapper].emit_op(Op::RETURN, line);
        let high = self.chunks[wrapper].scratch_high_water;
        self.chunks[wrapper].finalize_local_count(high);
        self.chunks[wrapper].direct_entry_has_captures = Some(false);
        Ok(wrapper)
    }

    /// `name(args)` straight to its typed entry. With `raw`, a native result
    /// stays a raw i64; otherwise it is boxed to the exact carrier.
    pub(super) fn try_emit_native_entry_call(
        &mut self,
        name: &str,
        args: &[Argument],
        raw: bool,
    ) -> Result<bool, String> {
        if self.native_function_entries.is_empty()
            || self.directives().declared_function_bindings_live != Some(false)
            || self.current_class.is_some()
            || self.has_accessible_local_binding(name)
            || args.iter().any(|arg| arg.by_ref || arg.spread || arg.name.is_some())
        {
            return Ok(false);
        }
        let Some(entry) = self.native_function_entries.get(self.canon_lookup(name).as_ref()).cloned() else {
            return Ok(false);
        };
        if entry.params.len() != args.len() || (raw && entry.result.is_none()) {
            return Ok(false);
        }
        let line = self.line;
        if self.chunks[entry.chunk].takes_receiver {
            crate::primitives::expressions::emit_undefined(self.chunk(), line);
        }
        for (arg, reading) in args.iter().zip(&entry.params) {
            if reading.is_some() {
                self.compile_exact_i64_operand(&arg.value)?;
            } else {
                self.compile_expr_with_value_copy(&arg.value)?;
            }
        }
        let mut signature = Chunk::new("<typed-source-function-entry>");
        signature.arity = self.chunks[entry.chunk].arity;
        signature.result_arity = 1;
        signature.func_sig = self.chunks[entry.chunk].func_sig.clone();
        signature.direct_entry_has_captures = Some(false);
        self.chunk().emit_direct_defined_call(entry.chunk as u32, &signature, false, line)?;
        if let (Some(reading), false) = (entry.result, raw) {
            let transport = self.directives().wide_integer_transport();
            crate::primitives::exact_integer::emit_box_i64_as(
                &mut self.chunks,
                self.current,
                reading == WordInterpretation::Signed,
                transport,
                line,
            );
        }
        Ok(true)
    }
}
