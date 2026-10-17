//! Optional result contracts for already-resolved compiler emitters.
//!
//! This sibling registration leaves `LanguageDef`, platform descriptors and
//! their legacy dispatch ABI unchanged. A consumer first resolves the normal
//! dispatcher, then asks for a matching typed descriptor. Names in source or
//! return-type hints do not establish a result-representation contract.

use std::sync::atomic::{AtomicUsize, Ordering};
use std::sync::{Mutex, OnceLock};

use super::EmitDispatchFn;
use crate::Chunk;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum EmitRequest {
    /// Preserve the language's ordinary value-position representation.
    Value,
    /// This call result is immediately discarded by its owning statement.
    DiscardedValue,
    /// The owning expression is directly consumed as an i32 condition.
    I32Condition,
    /// A statically selected memory32 integer operation consumes this result.
    I32Word,
    /// Preserve value-position representation, with an already-evaluated RHS
    /// physically produced by the compiler's string-literal emitter. The
    /// constant indexes its immutable string-global key in the current chunk.
    /// An inferred type or a mutable variable does not establish this fact.
    StringLiteralRhs { constant: u32 },
    /// RHS physically produced by the compiler's null-literal emitter.
    /// Declarations, variables and inferred nullable types do not prove this.
    NullLiteralRhs,
    /// At least one already-evaluated operand was physically emitted as a
    /// numeric, string, Boolean or null literal. Neither an annotation nor
    /// an inferred type establishes this fact. Preserve ordinary value output.
    ScalarLiteralOperand,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum ResultRepresentation {
    /// No representation proof; retain normal language truthiness handling.
    Unknown,
    /// The produced value is exactly `Value::Bool`, not a reference wrapper.
    BoxedBoolean,
    /// The produced value is an unboxed i32 Boolean (0 or 1).
    RawI32Boolean,
    /// An unboxed i32 bit pattern, not necessarily a Boolean or signed number.
    RawI32Word,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum EmitOutcome {
    /// Nothing was emitted. The consumer can continue its normal fallback.
    Unhandled,
    /// Emission succeeded; this describes the actual resulting stack value.
    Handled(ResultRepresentation),
}

pub type TypedEmitDispatchFn =
    fn(&str, &mut Vec<Chunk>, usize, u8, u32, EmitRequest) -> EmitOutcome;

/// A typed adaptation of one existing dispatcher, not a second name resolver.
///
/// The adaptation must retain its legacy dispatcher's supported operations.
/// It reports a representation only after successful emission. It must not
/// omit value-position Boolean materialization merely because it can produce
/// an i32 for a requested condition. `Unhandled` must not emit instructions.
#[derive(Clone, Copy)]
pub struct TypedEmitterDef {
    pub namespace: &'static str,
    pub legacy_dispatch: EmitDispatchFn,
    pub dispatch: TypedEmitDispatchFn,
}

fn registry() -> &'static Mutex<Vec<TypedEmitterDef>> {
    static REGISTRY: OnceLock<Mutex<Vec<TypedEmitterDef>>> = OnceLock::new();
    REGISTRY.get_or_init(|| Mutex::new(Vec::new()))
}

static GENERATION: AtomicUsize = AtomicUsize::new(0);

/// Register a typed adaptation without replacing an active legacy dispatcher.
/// Repeating the same descriptor is idempotent; conflicting registrations fail.
pub fn register(def: TypedEmitterDef) -> Result<(), String> {
    let mut entries = registry().lock().unwrap();
    if let Some(old) = entries.iter().find(|old| old.namespace == def.namespace) {
        if std::ptr::fn_addr_eq(old.legacy_dispatch, def.legacy_dispatch)
            && std::ptr::fn_addr_eq(old.dispatch, def.dispatch)
        {
            return Ok(());
        }
        return Err(format!(
            "typed emitter '{}' already has a different registration",
            def.namespace,
        ));
    }
    entries.push(def);
    GENERATION.fetch_add(1, Ordering::Release);
    Ok(())
}

/// Snapshot a descriptor only when it belongs to the selected dispatcher.
///
/// A descriptor for another implementation of the same namespace must not be
/// combined with first-registration-wins language/platform resolution. Return
/// `None` instead, retaining the selected legacy emitter's behavior. The caller
/// invokes the returned function only after this registry lock is released.
pub fn for_dispatch(
    namespace: &str,
    selected: EmitDispatchFn,
) -> Option<TypedEmitterDef> {
    if GENERATION.load(Ordering::Acquire) == 0 {
        return None;
    }
    registry()
        .lock()
        .unwrap()
        .iter()
        .find(|def| {
            def.namespace == namespace && std::ptr::fn_addr_eq(def.legacy_dispatch, selected)
        })
        .copied()
}

/// Generation for consumers that cache descriptor resolution. No cache here
/// stores function indexes, expression scopes, caller locations or VM values.
pub fn generation() -> usize {
    GENERATION.load(Ordering::Acquire)
}
