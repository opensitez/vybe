use vybe_compiler::primitives::namespaces::{self, NamespaceNode, Subtree};
use vybe_runtime::Value;

pub(crate) fn register_tree(root: &mut Subtree) {
    for (name, module, func) in [
        ("math.Sqrt", "ecma:math", "sqrt"),
        ("math.Abs", "ecma:math", "abs"),
        ("math.Floor", "ecma:math", "floor"),
        ("math.Ceil", "ecma:math", "ceil"),
        ("math.Min", "ecma:math", "min"),
        ("math.Max", "ecma:math", "max"),
        ("math.Pow", "ecma:math", "pow"),
        ("math.Sin", "ecma:math", "sin"),
        ("math.Cos", "ecma:math", "cos"),
        ("math.Tan", "ecma:math", "tan"),
        ("math.Asin", "ecma:math", "asin"),
        ("math.Acos", "ecma:math", "acos"),
        ("math.Atan", "ecma:math", "atan"),
        ("math.Atan2", "ecma:math", "atan2"),
        ("math.Log", "ecma:math", "log"),
        ("math.Exp", "ecma:math", "exp"),
        ("math.IsNaN", "ecma:number", "isNaN"),
    ] {
        insert_path(root, name, namespaces::host_fn(module, func));
    }

    for (name, emit) in [
        ("math.Round", "math.round_half_away"),
        ("math.NaN", "math.nan"),
        ("math.Inf", "math.inf"),
        ("math.IsInf", "math.is_inf"),
        ("math.Copysign", "math.copysign"),
        ("math.Signbit", "math.signbit"),
        ("math.Dim", "math.dim"),
        ("math.Float64bits", "bits.reinterpret_i64"),
        ("math.Float64frombits", "bits.reinterpret_f64"),
        ("math.Float32bits", "bits.reinterpret_i32"),
        ("math.Float32frombits", "bits.reinterpret_f32"),
        ("math.Mod", "expressions.f64_mod"),
        ("math.Remainder", "go.math_remainder"),
    ] {
        insert_path(root, name, NamespaceNode::CommonEmit(emit.to_string()));
    }

    for (name, value) in [
        ("math.Pi", std::f64::consts::PI),
        ("math.E", std::f64::consts::E),
        ("math.Ln2", std::f64::consts::LN_2),
        ("math.Ln10", std::f64::consts::LN_10),
        ("math.Sqrt2", std::f64::consts::SQRT_2),
        ("math.MaxFloat", f64::MAX),
        ("math.SmallestNonzeroFloat64", f64::MIN_POSITIVE),
    ] {
        insert_path(root, name, NamespaceNode::Const(Value::F64(value)));
    }
}

fn insert_path(root: &mut Subtree, path: &str, node: NamespaceNode) {
    let mut segments: Vec<&str> = path.split('.').collect();
    let Some(leaf) = segments.pop() else {
        return;
    };
    let mut cursor = root;
    for seg in segments {
        let entry = cursor
            .entry(seg.to_string())
            .or_insert_with(|| NamespaceNode::Namespace(Subtree::new()));
        let NamespaceNode::Namespace(children) = entry else {
            return;
        };
        cursor = children;
    }
    cursor.insert(leaf.to_string(), node);
}
