use vybe_runtime::{VM, Value};

#[test]
fn unformatted_holes_match_composite_format_and_evaluate_once() {
    vybe_language_csharp::register();
    let source = r#"
class Probe {
    public static int calls = 0;
    public static string Next() { calls = calls + 1; return "a"; }
    public static string Actual() {
        calls = 0;
        string missing = null;
        var value = $"{Next()}:{Next()}:{3.5}:{true}:{false}:{missing}:{{ok}}";
        return value + ":" + calls;
    }
    public static string Expected() {
        calls = 0;
        string missing = null;
        var value = String.Format("{0}:{1}:{2}:{3}:{4}:{5}:{{ok}}", Next(), Next(), 3.5, true, false, missing);
        return value + ":" + calls;
    }
}
"#;
    let module = vybe_language_csharp::parse(source).unwrap();
    let profile = vybe_compiler::profile::parse_profile(vybe_language_csharp::profile_source()).unwrap();
    let chunks = vybe_compiler::Compiler::with_profile(profile).compile(&module).unwrap();
    let actual = chunks.iter().find(|chunk| chunk.name == "Actual").unwrap();
    let expected = chunks.iter().find(|chunk| chunk.name == "Expected").unwrap();
    assert!(actual.code.len() < expected.code.len(), "constant interpolation must avoid the composite-format scanner");
    let mut vm = VM::new();
    vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);
    vm.run(chunks).unwrap();
    let class = vm.global("Probe").unwrap().clone();
    let actual = vm.resolve_property(&class, "Actual").unwrap();
    let expected = vm.resolve_property(&class, "Expected").unwrap();
    let actual = vm.invoke(&actual, &[]).unwrap();
    let expected = vm.invoke(&expected, &[]).unwrap();
    assert_eq!(actual, expected);
    let Value::String(text) = actual else { panic!("expected string") };
    assert!(text.ends_with(":2"), "each substitution must run exactly once");
}

#[test]
fn console_fixed_numeric_format_avoids_general_formatter() {
    vybe_language_csharp::register();
    let source = r#"using System;
Console.WriteLine(1.23456789123456);
Console.WriteLine(-0.0);
Console.WriteLine(0.0/0.0);
Console.WriteLine(1.0/0.0);
Console.WriteLine(-1.0/0.0);
Console.WriteLine(true);
Console.WriteLine((string)null);
Console.WriteLine($"value {3}");
"#;
    let module = vybe_language_csharp::parse(source).unwrap();
    let profile = vybe_compiler::profile::parse_profile(vybe_language_csharp::profile_source()).unwrap();
    let chunks = vybe_compiler::Compiler::with_profile(profile).compile(&module).unwrap();
    assert!(!chunks.iter().any(|chunk| chunk.name == "__stdlib_dotnet_numeric_format"), "fixed Console precision must not link the general formatter");
    let mut vm = VM::new();
    vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);
    let output = std::sync::Arc::new(std::sync::Mutex::new(Vec::new()));
    let capture = output.clone();
    vm.register_host_fn("wasi:cli/stdout", "write-via-stream", Box::new(move |ctx, args| {
        let (future, id) = ctx.create_future();
        let output = capture.clone();
        ctx.consume_byte_stream(&args[0], id, Box::new(move |bytes| {
            if let Some(bytes) = bytes { output.lock().unwrap().extend_from_slice(bytes); }
            Ok(())
        })).unwrap();
        future
    }));
    vm.run(chunks).unwrap();
    assert_eq!(&*output.lock().unwrap(), b"1.234567891235\n0\nNaN\ninf\n-inf\nTrue\n\nvalue 3\n");
}
