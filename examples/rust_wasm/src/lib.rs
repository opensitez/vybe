// Pure Rust functions compiled to WASM
// These can be called from VB/JS through the Vybe VM

#[no_mangle]
pub extern "C" fn add(a: i32, b: i32) -> i32 {
    a + b
}

#[no_mangle]
pub extern "C" fn factorial(n: i32) -> i32 {
    if n <= 1 { 1 } else { n * factorial(n - 1) }
}

#[no_mangle]
pub extern "C" fn fibonacci(n: i32) -> i32 {
    if n <= 1 { return n; }
    let mut a = 0;
    let mut b = 1;
    for _ in 2..=n {
        let c = a + b;
        a = b;
        b = c;
    }
    b
}

#[no_mangle]
pub extern "C" fn is_prime(n: i32) -> i32 {
    if n < 2 { return 0; }
    let mut i = 2;
    while i * i <= n {
        if n % i == 0 { return 0; }
        i += 1;
    }
    1
}

// String processing in linear memory
#[no_mangle]
pub extern "C" fn to_upper(ptr: i32, len: i32) -> i32 {
    let slice = unsafe { std::slice::from_raw_parts_mut(ptr as *mut u8, len as usize) };
    for byte in slice.iter_mut() {
        if *byte >= b'a' && *byte <= b'z' {
            *byte -= 32;
        }
    }
    ptr
}
