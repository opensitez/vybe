//! Installed Python relational ANY and native-join regression gates.
//! Builtin fallback arms materialize Bool; user/reflected dunder results remain ANY.
//! Source VM and complete native validation; native JS execution is a separate gate.
use vybe_compiler::primitives::platforms;
use vybe_runtime::{Chunk, VM, Value};
fn parse(_language: &str, src: &str) -> Vec<Chunk> {
    static ONCE: std::sync::Once = std::sync::Once::new();
    ONCE.call_once(vybe_language_python::register);
    let ast = vybe_language_python::parse(src).unwrap();
    let profile =
        vybe_compiler::profile::parse_profile(vybe_language_python::profile_source()).unwrap();
    vybe_compiler::Compiler::with_profile(profile)
        .compile(&ast)
        .unwrap()
}
fn global(vm: &VM, name: &str) -> Value {
    vm.global(name)
        .or_else(|| vm.global(&format!("${name}")))
        .unwrap_or_else(|| panic!("missing actual Source global {name}"))
        .clone()
}
fn compact(v: &Value) -> String {
    match v {
        Value::Object(o) => {
            let o = o.lock().unwrap();
            match &o.kind {
                vybe_runtime::value::ObjectKind::Array(items) => format!(
                    "Array({:?})",
                    items.iter().take(12).map(compact).collect::<Vec<_>>()
                ),
                _ => {
                    let scalar = |v: &Value| match v {
                        Value::Object(_) => "Object".to_string(),
                        _ => format!("{v:?}"),
                    };
                    format!(
                        "Object(type={}, fields={:?}, message={:?}, name={:?})",
                        o.type_id,
                        o.fields.iter().take(6).map(scalar).collect::<Vec<_>>(),
                        o.properties.get("message").map(scalar),
                        o.properties.get("name").map(scalar)
                    )
                }
            }
        }
        _ => format!("{v:?}"),
    }
}
fn execute(c: Vec<Chunk>) -> (VM, Result<Value, String>) {
    let bytes = c.iter().map(|c| c.code.len()).sum::<usize>();
    let locals = c.iter().map(|c| c.local_count as usize).sum::<usize>();
    let mut vm = VM::new();
    platforms::register_platforms_all(&mut vm);
    let result = vm.run(c).map_err(|mut e| {
        for f in &mut e.call_stack {
            f.offset = 0;
        }
        format!("{e:?}")
    });
    eprintln!(
        "actual Source complete bytecode={bytes} local_sum={locals} result={:?} last_exception={:?}",
        result.as_ref().map(compact),
        vm.last_exception.as_ref().map(compact)
    );
    for name in ["answer", "trace", "events", "calls", "escaped", "negated"] {
        let value = vm.global(name).or_else(|| vm.global(&format!("${name}")));
        eprintln!("actual Source {name}={:?}", value.map(compact));
    }
    (vm, result)
}
fn trace_items(vm: &VM) -> Vec<String> {
    let value = global(vm, "trace");
    let Value::Object(o) = value else {
        panic!("expected shared trace Array: {}", compact(&value));
    };
    let o = o.lock().unwrap();
    let vybe_runtime::value::ObjectKind::Array(items) = &o.kind else {
        panic!("expected trace Array");
    };
    items
        .iter()
        .map(|v| {
            let Value::String(s) = v else {
                panic!("trace item: {}", compact(v));
            };
            s.to_string()
        })
        .collect()
}
fn run(c: Vec<Chunk>) -> VM {
    let (vm, result) = execute(c);
    assert!(result.is_ok(), "actual Source result: {result:?}");
    assert!(
        vm.last_exception.is_none(),
        "actual Source uncaught exception: {:?}",
        vm.last_exception.as_ref().map(compact)
    );
    vm
}
const PYTHON_LT_ANY: &str = "trace = []\nclass Truth:\n    def __bool__(self):\n        trace.append('truth')\n        return False\ntruth = Truth()\nclass Compare:\n    def __lt__(self, other):\n        trace.append('compare')\n        return truth\nleft = Compare()\nright = Compare()\nanswer = left < right\nsame = answer is truth\nnegated = not answer\n";

#[test]
fn actual_python_relational_writable_dynlt_fallback_keeps_callback_and_boxing() {
    let source = "trace=[]\ndef __vybe_dynlt(a,b):\n    trace.append('fallback')\n    return 1\nanswer = 'a' < 'b'\n";
    let vm = run(parse("python", source));
    assert_eq!(global(&vm, "answer"), Value::Bool(true));
    assert_eq!(trace_items(&vm), ["fallback"]);
}

#[test]
fn strict_python_relational_any_dunder_result_tags_remain_values() {
    for (value, expected) in [
        ("False", Value::Bool(false)),
        ("7", Value::F64(7.0)),
        ("'sentinel'", Value::String("sentinel".into())),
        ("None", Value::Null),
    ] {
        let source = format!(
            "trace = []\ntarget = {value}\nclass Compare:\n    def __lt__(self, other):\n        trace.append('compare')\n        return target\nleft = Compare()\nright = Compare()\nanswer = left < right\n"
        );
        let vm = run(parse("python", &source));
        assert_eq!(
            global(&vm, "answer"),
            expected,
            "actual ANY return shape: {value}"
        );
        assert_eq!(trace_items(&vm), ["compare"]);
    }
}

#[test]
fn strict_python_reflected_notimplemented_any_priority_and_delayed_truth() {
    let source = "trace=[]\nclass Truth:\n    def __bool__(self):\n        trace.append('truth')\n        return False\ntruth=Truth()\nclass Left:\n    def __lt__(self,other):\n        trace.append('left')\n        return NotImplemented\nclass Right:\n    def __gt__(self,other):\n        trace.append('right')\n        return truth\nanswer=Left() < Right()\nnegated=not answer\n";
    let vm = run(parse("python", source));
    assert!(
        global(&vm, "answer") == global(&vm, "truth"),
        "reflected ANY identity"
    );
    assert_eq!(global(&vm, "negated"), Value::Bool(true));
    assert_eq!(trace_items(&vm), ["left", "right", "truth"]);
}

#[test]
fn strict_python_relational_truth_exception_occurs_after_comparison_once() {
    let source = "trace=[]\nclass Truth:\n    def __bool__(self):\n        trace.append('truth')\n        raise ValueError('truth failed')\ntruth=Truth()\nclass Compare:\n    def __lt__(self,other):\n        trace.append('compare')\n        return truth\ncaught=False\ntry:\n    answer=not (Compare() < Compare())\nexcept ValueError:\n    caught=True\n";
    let vm = run(parse("python", source));
    assert_eq!(global(&vm, "caught"), Value::Bool(true));
    assert_eq!(trace_items(&vm), ["compare", "truth"]);
}

#[test]
fn strict_python_rich_lt_any_value_identity_and_condition_truth_gate() {
    let vm = run(parse("python", PYTHON_LT_ANY));
    let answer = global(&vm, "answer");
    let truth = global(&vm, "truth");
    assert!(
        answer == truth,
        "actual ANY identity retained within VM: answer={} truth={}",
        compact(&answer),
        compact(&truth)
    );
    assert_eq!(global(&vm, "same"), Value::Bool(true));
    assert_eq!(global(&vm, "negated"), Value::Bool(true));
    assert_eq!(
        trace_items(&vm),
        ["compare", "truth"],
        "value comparison itself does not truth-test; later unary invokes exactly once"
    );
}

#[test]
fn actual_source_lt_route_is_explicit_registered_python_helper() {
    let ast = vybe_language_python::parse(PYTHON_LT_ANY).unwrap();
    let value = ast
        .body
        .iter()
        .find_map(|s| match &s.kind {
            vybe_ast::StmtKind::Assign { targets, value, .. }
                if targets
                    .iter()
                    .any(|t| matches!(&t.kind,vybe_ast::ExprKind::Ident(n) if n=="answer")) =>
            {
                Some(value)
            }
            _ => None,
        })
        .expect("actual answer assignment");
    assert!(
        matches!(&value.kind,vybe_ast::ExprKind::Call{callee,..} if matches!(&callee.kind,vybe_ast::ExprKind::Ident(n) if n=="__pylt__")),
        "actual LT AST: {value:?}"
    );
}
#[test]
fn strict_complete_actual_python_source_native_gate() {
    let c = parse("python", PYTHON_LT_ANY);
    let bytes = vybe_platform_wasm::write_wasm(&c);
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&bytes)
        .unwrap();
}

#[path = "fixtures/numeric_optimization_python_relational_pure.rs"]
mod pure_controls;
