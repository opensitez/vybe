/// Bytecode opcodes for the stack-based VM.
///
/// This is a language-agnostic instruction set. Language-specific semantics
/// (JS coercion, VB implicit conversion) are the compiler's responsibility —
/// the VM only executes typed operations.
///
/// Operand encoding:
/// - `u16` indices: two bytes big-endian after the opcode.
/// - `u8` counts: one byte after the opcode.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u8)]
pub enum Op {
    // -- Stack --
    /// Push constant from pool: [Const, hi, lo]
    Const,
    Pop,
    Dup,

    // -- Variables --
    GetLocal,   // u16 slot
    SetLocal,   // u16 slot
    GetGlobal,  // u16 name constant index
    SetGlobal,  // u16 name constant index
    GetUpvalue, // u8 index
    SetUpvalue, // u8 index

    // -- Properties --
    GetProp,   // u16 name constant index; stack [obj] → [val]
    SetProp,   // u16 name constant index; stack [obj, val] → [val]
    GetIndex,  // stack [obj, key] → [val]
    SetIndex,  // stack [obj, key, val] → [val]

    // -- Arithmetic (f64) --
    /// f64 + f64 → f64
    AddF,
    SubF,
    MulF,
    DivF,
    ModF,
    NegF,

    // -- Integer arithmetic (i32) --
    AddI,
    SubI,
    MulI,

    // -- String --
    /// String + String → String
    Concat,
    /// Concatenate N values from stack into one string: [StrConcat, u8 count]
    StrConcat,

    // -- Bitwise (i32) --
    BitAnd,
    BitOr,
    BitXor,
    BitNot,
    Shl,
    Shr,
    UShr,

    // -- Comparison --
    /// Same-type equality: push Bool
    CmpEq,
    CmpNe,
    /// Numeric comparison (f64): push Bool
    CmpLtF,
    CmpGtF,
    CmpLeF,
    CmpGeF,
    /// String comparison: push Bool
    CmpLtS,
    CmpGtS,

    // -- Logical (Bool operands only) --
    /// Bool → Bool
    BoolNot,

    // -- Control flow --
    /// Unconditional jump: [Jump, hi, lo] (signed i16 offset)
    Jump,
    /// Jump if TOS is Bool(false) (pops): [JumpIfFalse, hi, lo]
    JumpIfFalse,
    /// Jump if TOS is Bool(true) (pops): [JumpIfTrue, hi, lo]
    JumpIfTrue,
    /// Jump if TOS is Null (pops): [JumpIfNull, hi, lo]
    JumpIfNull,

    // -- Functions --
    /// Call bytecode function: [Call, u8 arg_count]
    Call,
    Return,
    /// Create closure: [Closure, u16 chunk_index, u8 upvalue_count, descriptors...]
    Closure,

    // -- Host functions --
    /// Call a host-registered function: [CallHost, u16 host_fn_index, u8 arg_count]
    /// Pops arg_count values, pushes one return value.
    CallHost,

    // -- Object / Array construction --
    NewObject,  // u16 property_count; stack [key, val, ...] → [obj]
    NewArray,   // u16 element_count; stack [elem, ...] → [arr]

    // -- Immediate values --
    PushNull,
    PushTrue,
    PushFalse,
    PushI32Zero,
    PushI32One,
    PushF64Zero,

    // -- Type checks (push Bool) --
    /// Is TOS Null?
    IsNull,
    /// Is TOS a String?
    IsString,
    /// Is TOS an F64?
    IsNumber,
    /// Is TOS a Bool?
    IsBool,
    /// Is TOS an Object (including Array, Function)?
    IsObject,
    /// Is TOS a Function?
    IsFunction,

    // -- Conversions (compiler emits these explicitly) --
    /// Coerce TOS to F64 (from I32/I64/Bool)
    ToF64,
    /// Coerce TOS to I32 (from F64/I64/Bool)
    ToI32,

    // -- Exception handling --
    TryStart,  // u16 catch_offset, u16 finally_offset
    TryEnd,
    Throw,

    // -- Iteration (future) --
    GetIterator,
    IterNext,
    Spread,

    // -- Class (future) --
    Class,     // u16 name
    Method,    // u16 name
    Inherit,

    Halt,
}

impl Op {
    pub fn from_byte(byte: u8) -> Option<Op> {
        if byte <= Op::Halt as u8 {
            Some(unsafe { std::mem::transmute(byte) })
        } else {
            None
        }
    }
}
