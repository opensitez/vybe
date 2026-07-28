//! Python source spelling -> shared protocol slot.
//!
//! Python-local by design: the shared class machinery sees only a
//! `SpecialMethodKind` (a numeric slot) and a canonical name. Which dunder
//! spells which role is Python's business and is decided here, before the
//! `NormalClass` reaches any shared code.

use vybe_bytecode::class_normalize::types::SpecialMethodKind;

/// Resolve a Python method name to `(canonical, slot?)`.
///
/// `None` means the method is ordinary — no protocol role.
pub fn canonical_method(name: &str) -> (String, Option<SpecialMethodKind>) {
    use SpecialMethodKind::*;

    match name {
        "__del__" => ("destructor".into(), Some(Destructor)),
        "__str__" => ("tostring".into(), Some(ToString)),
        "__repr__" => ("repr".into(), Some(Repr)),
        "__bool__" => ("bool".into(), Some(Bool)),
        // `__int__` and `__float__` are SEPARATE slots. They shared `ValueOf`
        // until 2026-07-28, so a class defining both published one method
        // under the other's slot.
        "__int__" => ("int".into(), Some(Int)),
        "__float__" => ("float".into(), Some(Float)),
        "__bytes__" => ("bytes".into(), Some(Bytes)),
        "__format__" => ("format".into(), Some(Format)),
        "__iter__" => ("iterator".into(), Some(Iterator)),
        "__next__" => ("next".into(), Some(Next)),
        "__aiter__" => ("asynciterator".into(), Some(AsyncIterator)),
        "__anext__" => ("asyncnext".into(), Some(AsyncNext)),
        "__reversed__" => ("reversed".into(), Some(Reversed)),
        "__add__" => ("add".into(), Some(Add)),
        "__sub__" => ("sub".into(), Some(Sub)),
        "__mul__" => ("mul".into(), Some(Mul)),
        "__truediv__" => ("div".into(), Some(Div)),
        // Was `Div` — the same slot as `__truediv__`, so `/` and `//` on one
        // class collided and the second install evicted the first.
        "__floordiv__" => ("floordiv".into(), Some(FloorDiv)),
        "__matmul__" => ("matmul".into(), Some(MatMul)),
        "__mod__" => ("mod".into(), Some(Mod)),
        "__pow__" => ("pow".into(), Some(Pow)),
        "__neg__" => ("neg".into(), Some(Neg)),
        "__pos__" => ("pos".into(), Some(Pos)),
        "__abs__" => ("abs".into(), Some(Abs)),
        "__round__" => ("round".into(), Some(Round)),
        "__floor__" => ("floor".into(), Some(Floor)),
        "__ceil__" => ("ceil".into(), Some(Ceil)),
        "__trunc__" => ("trunc".into(), Some(Trunc)),
        "__index__" => ("index".into(), Some(Index)),
        "__eq__" => ("eq".into(), Some(Eq)),
        "__ne__" => ("ne".into(), Some(Ne)),
        "__lt__" => ("lt".into(), Some(Lt)),
        "__le__" => ("le".into(), Some(Le)),
        "__gt__" => ("gt".into(), Some(Gt)),
        "__ge__" => ("ge".into(), Some(Ge)),
        "__and__" => ("and".into(), Some(And)),
        "__or__" => ("or".into(), Some(Or)),
        "__xor__" => ("xor".into(), Some(Xor)),
        "__invert__" => ("not".into(), Some(Not)),
        "__lshift__" => ("lshift".into(), Some(LShift)),
        "__rshift__" => ("rshift".into(), Some(RShift)),
        // Augmented assignment mutates in place — a different method from the
        // binary operator, so a different slot.
        "__iadd__" => ("iadd".into(), Some(IAdd)),
        "__isub__" => ("isub".into(), Some(ISub)),
        "__imul__" => ("imul".into(), Some(IMul)),
        "__itruediv__" => ("idiv".into(), Some(IDiv)),
        "__ifloordiv__" => ("ifloordiv".into(), Some(IFloorDiv)),
        "__imod__" => ("imod".into(), Some(IMod)),
        "__ipow__" => ("ipow".into(), Some(IPow)),
        "__imatmul__" => ("imatmul".into(), Some(IMatMul)),
        "__iand__" => ("iand".into(), Some(IAnd)),
        "__ior__" => ("ior".into(), Some(IOr)),
        "__ixor__" => ("ixor".into(), Some(IXor)),
        "__ilshift__" => ("ilshift".into(), Some(ILShift)),
        "__irshift__" => ("irshift".into(), Some(IRShift)),
        // Reflected: `2 + vec` dispatches onto the RIGHT operand.
        "__radd__" => ("radd".into(), Some(RAdd)),
        "__rsub__" => ("rsub".into(), Some(RSub)),
        "__rmul__" => ("rmul".into(), Some(RMul)),
        "__rtruediv__" => ("rdiv".into(), Some(RDiv)),
        "__rfloordiv__" => ("rfloordiv".into(), Some(RFloorDiv)),
        "__rmod__" => ("rmod".into(), Some(RMod)),
        "__rpow__" => ("rpow".into(), Some(RPow)),
        "__rmatmul__" => ("rmatmul".into(), Some(RMatMul)),
        "__rand__" => ("rand".into(), Some(RAnd)),
        "__ror__" => ("ror".into(), Some(ROr)),
        "__rxor__" => ("rxor".into(), Some(RXor)),
        "__rlshift__" => ("rlshift".into(), Some(RLShift)),
        "__rrshift__" => ("rrshift".into(), Some(RRShift)),
        "__len__" => ("len".into(), Some(Len)),
        "__getitem__" => ("getitem".into(), Some(GetItem)),
        "__setitem__" => ("setitem".into(), Some(SetItem)),
        "__delitem__" => ("delitem".into(), Some(DelItem)),
        "__missing__" => ("missing".into(), Some(Missing)),
        "__contains__" => ("contains".into(), Some(Contains)),
        "__call__" => ("call".into(), Some(Call)),
        "__instancecheck__" => ("hasinstance".into(), Some(HasInstance)),
        "__getattr__" | "__getattribute__" => ("getattr".into(), Some(GetAttr)),
        "__setattr__" => ("setattr".into(), Some(SetAttr)),
        "__delattr__" => ("delattr".into(), Some(DelAttr)),
        "__enter__" => ("enter".into(), Some(Enter)),
        "__exit__" => ("exit".into(), Some(Exit)),
        "__aenter__" => ("asyncenter".into(), Some(AsyncEnter)),
        "__aexit__" => ("asyncexit".into(), Some(AsyncExit)),
        "__copy__" => ("clone".into(), Some(Clone)),
        "__getstate__" => ("serialize".into(), Some(Serialize)),
        "__setstate__" => ("deserialize".into(), Some(Deserialize)),
        "__hash__" => ("hash".into(), Some(Hash)),
        _ => (name.to_string(), None),
    }
}
