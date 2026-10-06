//! Registered number radix checks without guest compiler dependencies.
use vybe_runtime::value::{Object, ObjectKind, Value};
use vybe_runtime::VM;

fn original_radix(number: i64, radix: u32) -> String {
    let mut value = (number as i128).unsigned_abs();
    if value == 0 { return "0".to_owned(); }
    let mut result = String::new();
    while value != 0 {
        result.insert(0, char::from_digit((value % radix as u128) as u32, radix).unwrap());
        value /= radix as u128;
    }
    if number < 0 { result.insert(0, '-'); }
    result
}

fn main() {
    let mut vm = VM::new();
    vybe_platform_ecma::number::register(&mut vm);
    let index = *vm.host_registry.get(&("ecma:number".to_owned(), "toString".to_owned())).unwrap();
    let mut object = Object::new();
    object.kind = ObjectKind::HostFunction(index);
    let callable = Value::Object(vybe_runtime::heap::alloc(object));
    for number in [0, 1, -1, 42, -42, i32::MAX as i64, i32::MIN as i64, i64::MAX, i64::MIN] {
        for radix in 2..=36 {
            if radix == 10 { continue; }
            // Preserve the registered adapter's existing f64-to-i64 conversion.
            let converted = (number as f64) as i64;
            let expected = original_radix(converted, radix);
            let result = vm.invoke(&callable, &[Value::F64(number as f64), Value::I32(radix as i32)]).unwrap();
            let Value::String(text) = result else { panic!("string expected"); };
            assert_eq!(text.as_ref(), expected, "number={number}, radix={radix}");
        }
    }
    println!("Registered integer radix checks passed for all nondecimal radices and signed limits");
}
