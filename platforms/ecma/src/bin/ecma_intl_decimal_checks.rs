//! Public registered NumberFormat checks; no test-only library APIs.
use std::hint::black_box;
use std::time::Instant;
use icu::decimal::{DecimalFormatter, input::Decimal};
use vybe_runtime::value::{Object, ObjectKind};
use vybe_runtime::{VM, Value};

fn string(value: &str) -> Value { Value::String(value.into()) }

fn function(vm: &VM, name: &str) -> Value {
    let index = vm.resolve_host_function_index("ecma:intl/numberformat", name).unwrap();
    let mut object = Object::new();
    object.kind = ObjectKind::HostFunction(index);
    Value::Object(vybe_runtime::heap::alloc(object))
}

fn text(value: Value) -> String {
    match value {
        Value::String(value) => value.to_string(),
        other => panic!("expected formatted string, got {other:?}"),
    }
}

fn set(formatter: &Value, key: &str, value: Value) {
    let Value::Object(object) = formatter else { panic!("formatter object missing") };
    object.lock().unwrap().properties.insert(key.into(), value);
}

fn main() {
    let mut vm = VM::new();
    vybe_platform_ecma::intl::register(&mut vm);
    let constructor = function(&vm, "new");
    let format = function(&vm, "format");
    for locale in ["en-US", "fr-FR", "de-DE", "ar", "hi-IN", "ja", "zh", "es", "it", "nl"] {
        let formatter = vm.invoke(&constructor, &[string(locale)]).unwrap();
        let icu_locale: icu::locale::Locale = locale.parse().unwrap();
        let reference = DecimalFormatter::try_new((&icu_locale).into(), Default::default()).unwrap();
        for number in [-1234_i64, 0, 1, 1234567] {
            let actual = text(vm.invoke(&format, &[formatter.clone(), Value::F64(number as f64)]).unwrap());
            let expected = reference.format(&Decimal::from(number)).to_string();
            assert_eq!(actual, expected, "locale={locale}, number={number}");
        }
    }
    let formatter = vm.invoke(&constructor, &[string("en-US")]).unwrap();
    let args = [formatter.clone(), Value::F64(1234.5)];
    assert_eq!(text(vm.invoke(&format, &args).unwrap()), "1,234.5");
    set(&formatter, "useGrouping", Value::Bool(false));
    assert_eq!(text(vm.invoke(&format, &args).unwrap()), "1234.5");
    set(&formatter, "useGrouping", Value::Bool(true));
    set(&formatter, "locale", string("de-DE"));
    assert_eq!(text(vm.invoke(&format, &args).unwrap()), "1.234,5");
    set(&formatter, "locale", string("en-US"));
    assert_eq!(text(vm.invoke(&format, &args).unwrap()), "1,234.5");
    println!("registered NumberFormat ICU equivalence across ten locales and fresh locale/option reads passed");
    let start = Instant::now();
    for _ in 0..10000 { black_box(vm.invoke(&format, black_box(&args)).unwrap()); }
    println!("10000 warm registered NumberFormat calls (VM dispatch included, setup excluded): {:?}", start.elapsed());
}
