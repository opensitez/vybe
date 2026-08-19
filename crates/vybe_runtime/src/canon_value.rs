//! Canonical ABI value transfer — `CanonicalABI.md` §`store` / §`load`.
//!
//! Moves ONE value of a component type between a `Value` and linear memory,
//! using the layout in [`crate::canon_layout`]. Together they are what
//! `canon future.{read,write}` needs: that built-in is
//! `(func (param i32 T) (result i32))` — a handle and a pointer, with **no
//! count** — because a future carries exactly one element whose size and
//! encoding come from its type.
//!
//! Both entry points begin with the two assertions the spec opens with:
//!
//! ```python
//! assert(ptr == align_to(ptr, alignment(t, ...)))
//! assert(ptr + elem_size(t, ...) <= len(cx.opts.memory))
//! ```
//!
//! Here they are real errors rather than debug assertions, because they are
//! reachable from guest code: a component that hands over a misaligned pointer
//! must be told, not silently allowed to write across a field boundary.
//!
//! **Not every `ValType` is covered**, and the uncovered ones return an error
//! naming themselves rather than moving an approximate number of bytes. A
//! canonical ABI that is quietly wrong about layout is worse than one that
//! refuses: the first corrupts a peer's memory, the second fails a test.

use crate::canon_layout::{align_to, alignment, elem_size};
use crate::component::ValType;
use crate::value::Value;

/// Why a canonical transfer could not be performed.
#[derive(Debug, Clone)]
pub enum CanonError {
    Misaligned {
        ptr: u32,
        required: u32,
    },
    OutOfBounds {
        ptr: u32,
        size: u32,
        memory_len: usize,
    },
    /// The type has a canonical layout this crate has not implemented yet.
    /// Named so the message points at the work rather than at the symptom.
    Unsupported(&'static str),
}

impl std::fmt::Display for CanonError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            CanonError::Misaligned { ptr, required } => write!(
                f,
                "canonical ABI: pointer {ptr} is not aligned to {required} bytes"
            ),
            CanonError::OutOfBounds {
                ptr,
                size,
                memory_len,
            } => write!(
                f,
                "canonical ABI: {size} bytes at {ptr} is out of bounds of linear memory ({memory_len} bytes)"
            ),
            CanonError::Unsupported(what) => write!(
                f,
                "canonical ABI: no store/load implemented for {what} — see CanonicalABI.md §Storing/§Loading"
            ),
        }
    }
}

/// The two guards every `store`/`load` opens with.
fn check(ptr: u32, t: &ValType, memory_len: usize) -> Result<u32, CanonError> {
    let required = alignment(t);
    if ptr != align_to(ptr, required) {
        return Err(CanonError::Misaligned { ptr, required });
    }
    let size = elem_size(t);
    if (ptr as usize).saturating_add(size as usize) > memory_len {
        return Err(CanonError::OutOfBounds {
            ptr,
            size,
            memory_len,
        });
    }
    Ok(size)
}

/// `store(cx, v, t, ptr)` — write one value of type `t` at `ptr`.
pub fn store(
    memory: &crate::shared_memory::SharedMemory,
    v: &Value,
    t: &ValType,
    ptr: u32,
) -> Result<(), CanonError> {
    check(ptr, t, memory.len())?;
    let addr = ptr as usize;
    match t {
        // `store_int(cx, int(bool(v)), ptr, 1)` — a bool is one byte, 0 or 1,
        // never the raw truthiness of whatever the guest had.
        ValType::Bool => {
            let _ = memory.store_u8(addr, u8::from(v.as_bool()));
        }
        ValType::I32 => {
            let _ = memory.store_i32(addr, v.as_i32());
        }
        ValType::I64 => {
            let _ = memory.store_i64(addr, v.as_i64());
        }
        ValType::F64 => {
            let _ = memory.store_f64(addr, v.as_f64());
        }
        // A handle is its i32 index into the handle table.
        ValType::Own(_) | ValType::Borrow(_) | ValType::Stream(_) | ValType::Future(_) => {
            let _ = memory.store_i32(addr, v.as_i32());
        }
        // `string` is (ptr, length) — storing one means allocating the bytes
        // somewhere, which is `realloc`'s job and needs the canonical options
        // this crate does not carry yet.
        ValType::String => return Err(CanonError::Unsupported("string (needs realloc)")),
        ValType::List(_) => return Err(CanonError::Unsupported("list (needs realloc)")),
        ValType::Record(_) => return Err(CanonError::Unsupported("record")),
        ValType::Option(_) => return Err(CanonError::Unsupported("option")),
        ValType::Result(_, _) => return Err(CanonError::Unsupported("result")),
        ValType::Any => return Err(CanonError::Unsupported("any (not a component type)")),
    }
    Ok(())
}

/// `load(cx, ptr, t)` — read one value of type `t` from `ptr`.
pub fn load(
    memory: &crate::shared_memory::SharedMemory,
    t: &ValType,
    ptr: u32,
) -> Result<Value, CanonError> {
    check(ptr, t, memory.len())?;
    let addr = ptr as usize;
    Ok(match t {
        // `convert_int_to_bool` — any non-zero byte is true.
        ValType::Bool => Value::Bool(memory.load_u8(addr).unwrap_or(0) != 0),
        ValType::I32 => Value::I32(memory.load_i32(addr).unwrap_or(0)),
        ValType::I64 => Value::I64(memory.load_i64(addr).unwrap_or(0)),
        ValType::F64 => Value::F64(memory.load_f64(addr).unwrap_or(0.0)),
        ValType::Own(_) | ValType::Borrow(_) | ValType::Stream(_) | ValType::Future(_) => {
            Value::I32(memory.load_i32(addr).unwrap_or(0))
        }
        ValType::String => return Err(CanonError::Unsupported("string")),
        ValType::List(_) => return Err(CanonError::Unsupported("list")),
        ValType::Record(_) => return Err(CanonError::Unsupported("record")),
        ValType::Option(_) => return Err(CanonError::Unsupported("option")),
        ValType::Result(_, _) => return Err(CanonError::Unsupported("result")),
        ValType::Any => return Err(CanonError::Unsupported("any (not a component type)")),
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::shared_memory::SharedMemory;

    fn mem() -> SharedMemory {
        SharedMemory::new(1024)
    }

    #[test]
    fn scalars_round_trip() {
        let m = mem();
        store(&m, &Value::I32(0x1234_5678), &ValType::I32, 0).unwrap();
        assert_eq!(load(&m, &ValType::I32, 0).unwrap().as_i32(), 0x1234_5678);

        store(&m, &Value::I64(-9), &ValType::I64, 8).unwrap();
        assert_eq!(load(&m, &ValType::I64, 8).unwrap().as_i64(), -9);

        store(&m, &Value::F64(1.5), &ValType::F64, 16).unwrap();
        assert_eq!(load(&m, &ValType::F64, 16).unwrap().as_f64(), 1.5);
    }

    #[test]
    fn bool_stores_one_byte_not_truthiness() {
        let m = mem();
        store(&m, &Value::Bool(true), &ValType::Bool, 0).unwrap();
        assert_eq!(m.load_u8(0).unwrap(), 1);
        // and any non-zero byte reads back as true
        m.store_u8(1, 0x7f).unwrap();
        assert!(load(&m, &ValType::Bool, 1).unwrap().as_bool());
    }

    #[test]
    fn misaligned_pointer_is_an_error_not_a_silent_write() {
        let m = mem();
        // i64 needs 8-byte alignment; 4 is not.
        let e = store(&m, &Value::I64(1), &ValType::I64, 4).unwrap_err();
        assert!(matches!(e, CanonError::Misaligned { required: 8, .. }));
    }

    #[test]
    fn out_of_bounds_is_an_error() {
        let m = SharedMemory::new(16);
        let e = store(&m, &Value::I64(1), &ValType::I64, 16).unwrap_err();
        assert!(matches!(e, CanonError::OutOfBounds { .. }));
    }

    #[test]
    fn unimplemented_types_refuse_rather_than_guess() {
        let m = mem();
        // The point: a wrong number of bytes would corrupt a peer's memory
        // silently. Refusing fails a test instead.
        assert!(matches!(
            load(&m, &ValType::String, 0),
            Err(CanonError::Unsupported(_))
        ));
    }
}
