//! Arrays of scalars stored as Wasm GC `(array (mut <element>))`.
//!
//! A language declares the four operations of each element kind as builtins
//! bound to the `collections.<kind>_array_{new,get,set,len}` emitters. The
//! index is an i32 and the count an i64 word; an element is a raw word of its
//! storage type, converted from and to the language value only where a boxed
//! position reads or writes it. Each array access records its type row in
//! `Chunk::typed_array_ops`, which the Wasm writer reads to emit
//! `array.get`/`array.set`/`array.len` on the declared array type.
use super::*;

/// The `collections.<kind>_array_*` operation a builtin is bound to.
#[derive(Clone, Copy, PartialEq, Eq)]
pub(super) enum TypedArrayOp {
    /// `(length)` → a zero-filled array.
    New,
    /// `(array, index)` → the element; `(array, index, failure)` first checks
    /// `index` against the length and evaluates `failure` when it is out of
    /// range (the operands are side-effect free and read twice).
    Get,
    /// `(array, index, value)` → stores the element.
    Set,
    /// `(array)` → the element count.
    Len,
}

/// The element kind of a typed array: its storage type and the value an
/// element reads as.
#[derive(Clone, Copy, PartialEq, Eq)]
pub(super) enum TypedElement {
    /// i64 storage; an exact signed 64-bit integer value.
    I64,
    /// i32 storage of 0/1; a Boolean value.
    Bool,
    /// i32 storage; a signed integer Number.
    I32,
    /// i32 storage of the bits; an unsigned integer Number.
    U32,
    /// f64 storage; a Number.
    F64,
    /// f32 storage; a Number rounded to single precision.
    F32,
}

impl TypedElement {
    fn from_name(name: &str) -> Option<Self> {
        Some(match name {
            "i64" => Self::I64,
            "bool" => Self::Bool,
            "i32" => Self::I32,
            "u32" => Self::U32,
            "f64" => Self::F64,
            "f32" => Self::F32,
            _ => return None,
        })
    }

    /// The storage type, which names the type-table row.
    fn storage(self) -> &'static str {
        match self {
            Self::I64 => "i64",
            Self::Bool | Self::I32 | Self::U32 => "i32",
            Self::F64 => "f64",
            Self::F32 => "f32",
        }
    }
}

impl Compiler {
    /// The typed-array operation `callee` names, if it is a builtin bound to one.
    pub(super) fn typed_array_op(&self, callee: &str) -> Option<(TypedArrayOp, TypedElement)> {
        let def = self.profile.lookup_builtin(callee)?;
        let vybe_runtime::profile::BuiltinEmit::Common(emit) = &def.emit else { return None };
        let (kind, op) = emit.strip_prefix("collections.")?.split_once("_array_")?;
        let op = match op {
            "new" => TypedArrayOp::New,
            "get" => TypedArrayOp::Get,
            "set" => TypedArrayOp::Set,
            "len" => TypedArrayOp::Len,
            _ => return None,
        };
        Some((op, TypedElement::from_name(kind)?))
    }

    /// A call whose raw result is an i64 word: an `i64` element read or a count.
    pub(super) fn typed_array_i64_word(&self, callee: &Expression) -> Option<(TypedArrayOp, TypedElement)> {
        let ExprKind::Ident(name) = &callee.kind else { return None };
        self.typed_array_op(name).filter(|(op, element)| {
            *op == TypedArrayOp::Len || (*op == TypedArrayOp::Get && *element == TypedElement::I64)
        })
    }

    /// A read of a `bool` element, whose raw word is an i32 condition.
    pub(super) fn typed_array_bool_read(&self, value: &Expression) -> bool {
        let ExprKind::Call { callee, optional: false, .. } = &value.kind else { return false };
        let ExprKind::Ident(name) = &callee.kind else { return false };
        self.typed_array_op(name) == Some((TypedArrayOp::Get, TypedElement::Bool))
    }

    /// The 1-based type-table index of the array row for `element`.
    fn typed_array_row(&mut self, element: TypedElement) -> u16 {
        let storage = element.storage();
        let (name, field) = vybe_runtime::chunk::TYPED_ARRAY_ROWS
            .iter()
            .copied()
            .find(|(_, field)| *field == storage)
            .expect("every storage type has a typed array row");
        let index = match self.chunks[0].types.iter().position(|row| row.name == name) {
            Some(position) => position + 1,
            None => crate::primitives::classes::register_gc_array_type(&mut self.chunks, name, field),
        };
        u16::try_from(index).expect("type table index fits u16")
    }

    /// Emit `op` on the typed array, recording its row for the writer.
    fn emit_typed_array_op(&mut self, op: Op, element: TypedElement) {
        let row = self.typed_array_row(element);
        let at = self.chunk().code.len();
        self.chunk().typed_array_ops.insert(at, row);
        self.emit(op);
    }

    /// The i32 index operand of an access.
    fn compile_typed_array_index(&mut self, index: &Expression) -> Result<(), String> {
        self.compile_exact_i64_operand(index)?;
        self.emit(Op::I32_WRAP_I64);
        Ok(())
    }

    /// The raw storage word of an element value.
    fn compile_typed_element(&mut self, element: TypedElement, value: &Expression) -> Result<(), String> {
        match element {
            TypedElement::I64 => self.compile_exact_i64_operand(value)?,
            TypedElement::Bool => self.compile_condition_to_i32(value)?,
            TypedElement::I32 | TypedElement::U32 | TypedElement::F64 | TypedElement::F32 => {
                self.compile_expr(value)?;
                let to_f64 = self.import("wasm:js-number", "toF64");
                self.emit_host_call(to_f64, 1);
                match element {
                    TypedElement::I32 => self.emit(Op::I32_TRUNC_SAT_F64_S),
                    TypedElement::U32 => self.emit(Op::I32_TRUNC_SAT_F64_U),
                    TypedElement::F32 => self.emit(Op::F32_DEMOTE_F64),
                    _ => {}
                }
            }
        }
        Ok(())
    }

    /// Compile `op` leaving its raw result: the array reference for `New`, the
    /// element's storage word for `Get`, an i64 count for `Len`, nothing for
    /// `Set`.
    pub(super) fn compile_typed_array_word(
        &mut self,
        op: TypedArrayOp,
        element: TypedElement,
        args: &[&Expression],
    ) -> Result<(), String> {
        let line = self.line;
        let arity_ok = match op {
            TypedArrayOp::New | TypedArrayOp::Len => args.len() == 1,
            TypedArrayOp::Get => matches!(args.len(), 2 | 3),
            TypedArrayOp::Set => args.len() == 3,
        };
        if !arity_ok {
            return Err(format!("typed array operation has {} arguments", args.len()));
        }
        match op {
            TypedArrayOp::New => {
                self.compile_typed_array_index(args[0])?;
                let row = self.typed_array_row(element);
                self.chunk().emit_op_u16(Op::ARRAY_NEW_DEFAULT, row, line);
            }
            TypedArrayOp::Get => {
                if let Some(failure) = args.get(2) {
                    self.compile_exact_i64_operand(args[1])?;
                    self.compile_expr(args[0])?;
                    self.emit_typed_array_op(Op::ARRAY_LENGTH, element);
                    self.emit(Op::I64_EXTEND_I32_U);
                    self.emit(Op::I64_LT_U);
                    self.emit(Op::I32_EQZ);
                    self.chunk().emit_if(line);
                    self.compile_expr(failure)?;
                    self.emit(Op::DROP);
                    self.chunk().emit_end(line);
                }
                self.compile_expr(args[0])?;
                self.compile_typed_array_index(args[1])?;
                self.emit_typed_array_op(Op::ARRAY_GET, element);
            }
            TypedArrayOp::Set => {
                self.compile_expr(args[0])?;
                self.compile_typed_array_index(args[1])?;
                self.compile_typed_element(element, args[2])?;
                self.emit_typed_array_op(Op::ARRAY_SET, element);
            }
            TypedArrayOp::Len => {
                self.compile_expr(args[0])?;
                self.emit_typed_array_op(Op::ARRAY_LENGTH, element);
                self.emit(Op::I64_EXTEND_I32_U);
            }
        }
        Ok(())
    }

    /// Compile `op` in a value position: an element or a count is boxed as
    /// the language value of its kind, and a store yields no value.
    pub(super) fn compile_typed_array_value(
        &mut self,
        op: TypedArrayOp,
        element: TypedElement,
        args: &[&Expression],
    ) -> Result<(), String> {
        let line = self.line;
        self.compile_typed_array_word(op, element, args)?;
        match (op, element) {
            (TypedArrayOp::New, _) => {}
            (TypedArrayOp::Set, _) => self.emit(Op::NULL),
            (TypedArrayOp::Len, _) | (TypedArrayOp::Get, TypedElement::I64) => {
                crate::primitives::exact_integer::emit_box_i64(&mut self.chunks, self.current, true, line);
            }
            (TypedArrayOp::Get, TypedElement::Bool) => crate::primitives::ops::emit_i32_to_bool(self.chunk(), line),
            (TypedArrayOp::Get, TypedElement::I32) => {
                let from_i32 = self.import("wasm:js-number", "fromI32");
                self.emit_host_call(from_i32, 1);
            }
            (TypedArrayOp::Get, TypedElement::U32 | TypedElement::F64 | TypedElement::F32) => {
                match element {
                    TypedElement::U32 => self.emit(Op::F64_CONVERT_I32_U),
                    TypedElement::F32 => self.emit(Op::F64_PROMOTE_F32),
                    _ => {}
                }
                let from_f64 = self.import("wasm:js-number", "fromF64");
                self.emit_host_call(from_f64, 1);
            }
        }
        Ok(())
    }
}
