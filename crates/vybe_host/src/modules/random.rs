use std::cell::RefCell;
use std::rc::Rc;
use vybe_bytecode::{VM, Value};
use vybe_bytecode::value::Object;

// Simple xorshift64 PRNG state — thread-local for safety.
thread_local! {
    static RNG_STATE: RefCell<u64> = RefCell::new(
        std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap_or_default()
            .as_nanos() as u64
    );
}

fn next_u64() -> u64 {
    RNG_STATE.with(|state| {
        let mut s = state.borrow_mut();
        // xorshift64
        *s ^= *s << 13;
        *s ^= *s >> 7;
        *s ^= *s << 17;
        *s
    })
}

fn next_f64() -> f64 {
    (next_u64() >> 11) as f64 / (1u64 << 53) as f64
}

pub fn register(vm: &mut VM) {
    // Random float in [0, 1)
    vm.register_host_fn("wasi:random", "random", Box::new(|_vm: &mut VM, _args: &[Value]| {
        Value::F64(next_f64())
    }));

    // Random integer in [min, max] inclusive
    vm.register_host_fn("wasi:random", "randomInt", Box::new(|_vm: &mut VM, args: &[Value]| {
        let min = args.first().map(|v| v.as_f64() as i64).unwrap_or(0);
        let max = args.get(1).map(|v| v.as_f64() as i64).unwrap_or(100);
        if max <= min { return Value::F64(min as f64); }
        let range = (max - min + 1) as u64;
        let val = min + (next_u64() % range) as i64;
        Value::F64(val as f64)
    }));

    // Seed the RNG
    vm.register_host_fn("wasi:random", "seed", Box::new(|_vm: &mut VM, args: &[Value]| {
        let s = args.first().map(|v| v.as_f64() as u64).unwrap_or(42);
        RNG_STATE.with(|state| *state.borrow_mut() = s);
        Value::Null
    }));

    // Random bytes (returns array of u8 values)
    vm.register_host_fn("wasi:random", "randomBytes", Box::new(|_vm: &mut VM, args: &[Value]| {
        let n = args.first().map(|v| v.as_f64() as usize).unwrap_or(16);
        let mut bytes = Vec::with_capacity(n);
        for _ in 0..n {
            bytes.push(Value::F64((next_u64() & 0xFF) as f64));
        }
        Value::Object(Rc::new(RefCell::new(Object::new_array(bytes))))
    }));

    // Random UUID v4
    vm.register_host_fn("wasi:random", "uuid", Box::new(|_vm: &mut VM, _args: &[Value]| {
        let a = next_u64();
        let b = next_u64();
        let s = format!(
            "{:08x}-{:04x}-4{:03x}-{:04x}-{:012x}",
            (a >> 32) as u32,
            (a >> 16) as u16 & 0xFFFF,
            a as u16 & 0x0FFF,
            (b >> 48) as u16 & 0x3FFF | 0x8000,
            b & 0xFFFFFFFFFFFF,
        );
        Value::String(Rc::from(s.as_str()))
    }));
}
