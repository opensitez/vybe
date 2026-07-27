use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: Class Expressions (Named & Anonymous Expressions)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_anonymous_class_expression_instantiation() {
    let src = r#"
const MyClass = class {
    greet() { return "Hello from Anon Class"; }
};
console.log(new MyClass().greet());
"#;
    assert_eq!(run_js(src), vec!["Hello from Anon Class"]);
}

#[test]
fn test_js_anonymous_class_expression_name_inference() {
    let src = r#"
const Widget = class {};
console.log(Widget.name);
"#;
    assert_eq!(run_js(src), vec!["Widget"]);
}

#[test]
fn test_js_named_class_expression_name_property() {
    let src = r#"
const Foo = class Bar {};
console.log(Foo.name);
"#;
    assert_eq!(run_js(src), vec!["Bar"]);
}

#[test]
fn test_js_named_class_expression_internal_name_binding() {
    let src = r#"
const FactorialClass = class Fact {
    static calc(n) {
        if (n <= 1) return 1;
        return n * Fact.calc(n - 1); // Fact internal binding accessible inside class body!
    }
};
console.log(FactorialClass.calc(5));
"#;
    assert_eq!(run_js(src), vec!["120"]);
}

#[test]
fn test_js_named_class_expression_internal_name_not_accessible_outside() {
    let src = r#"
const Container = class SecretClass {};
try {
    eval("SecretClass");
} catch (e) {
    console.log("Outside Named Class Expression ReferenceError");
}
"#;
    assert_eq!(
        run_js(src),
        vec!["Outside Named Class Expression ReferenceError"]
    );
}

#[test]
fn test_js_class_expression_passed_as_function_argument() {
    let src = r#"
function instantiate(ClassRef, arg) {
    return new ClassRef(arg);
}
const item = instantiate(class {
    constructor(v) { this.v = v; }
}, "ArgumentVal");
console.log(item.v);
"#;
    assert_eq!(run_js(src), vec!["ArgumentVal"]);
}

#[test]
fn test_js_class_expression_immediately_invoked() {
    let src = r#"
const singleton = new (class {
    constructor() { this.version = "1.0.0"; }
})();
console.log(singleton.version);
"#;
    assert_eq!(run_js(src), vec!["1.0.0"]);
}

#[test]
fn test_js_class_expression_in_object_literal_property() {
    let src = r#"
const factory = {
    UserClass: class {
        constructor(name) { this.name = name; }
    }
};
const u = new factory.UserClass("Bob");
console.log(u.name);
"#;
    assert_eq!(run_js(src), vec!["Bob"]);
}

#[test]
fn test_js_class_expression_extends_clause() {
    let src = r#"
class Base {
    static baseValue = 10;
}
const DerivedClass = class extends Base {
    static getDouble() { return super.baseValue * 2; }
};
console.log(DerivedClass.getDouble());
"#;
    assert_eq!(run_js(src), vec!["20"]);
}

#[test]
fn test_js_class_expression_private_fields_and_methods() {
    let src = r#"
const Vault = class {
    #secret = "Pass";
    #getSecret() { return this.#secret; }
    reveal() { return this.#getSecret(); }
};
console.log(new Vault().reveal());
"#;
    assert_eq!(run_js(src), vec!["Pass"]);
}

#[test]
fn test_js_class_expression_static_initialization_block() {
    let src = r#"
const App = class {
    static status;
    static {
        this.status = "Ready";
    }
};
console.log(App.status);
"#;
    assert_eq!(run_js(src), vec!["Ready"]);
}

#[test]
fn test_js_class_expression_super_access_inside_static_initializer_and_block() {
    let src = r#"
class Base {
    static baseValue = 10;
}

const Derived = class extends Base {
    static fromBase = super.baseValue;
    static fromBaseAndBlock;
    static {
        this.fromBaseAndBlock = super.baseValue + this.fromBase;
    }
};

console.log(Derived.fromBase);
console.log(Derived.fromBaseAndBlock);
"#;
    assert_eq!(run_js(src), vec!["10", "20"]);
}

#[test]
fn test_js_named_class_expression_internal_name_immutable() {
    let src = r#"
const Foo = class Bar {
    static tryRebind() {
        "use strict";
        try {
            Bar = 123; // Internal class expression binding is read-only constant!
        } catch (e) {
            console.log("Internal Name Immutable TypeError");
        }
    }
};
Foo.tryRebind();
"#;
    assert_eq!(run_js(src), vec!["Internal Name Immutable TypeError"]);
}

#[test]
fn test_js_class_expression_in_array_elements() {
    let src = r#"
const classList = [
    class { getType() { return "TypeA"; } },
    class { getType() { return "TypeB"; } }
];
console.log(new classList[0]().getType() + "|" + new classList[1]().getType());
"#;
    assert_eq!(run_js(src), vec!["TypeA|TypeB"]);
}

#[test]
fn test_js_class_expression_return_from_factory_function() {
    let src = r#"
function createModel(modelName) {
    return class {
        static name = modelName;
        getModelName() { return modelName; }
    };
}
const UserModel = createModel("User");
console.log(new UserModel().getModelName());
"#;
    assert_eq!(run_js(src), vec!["User"]);
}

#[test]
fn test_js_class_expression_computed_method_names() {
    let src = r#"
const methodName = "dynamicExec";
const DynamicClass = class {
    [methodName]() { return "ExecSuccess"; }
};
console.log(new DynamicClass().dynamicExec());
"#;
    assert_eq!(run_js(src), vec!["ExecSuccess"]);
}

#[test]
fn test_js_class_expression_generator_methods() {
    let src = r#"
const GenClass = class {
    *items() {
        yield 1; yield 2;
    }
};
console.log([...new GenClass().items()].join(","));
"#;
    assert_eq!(run_js(src), vec!["1,2"]);
}

#[test]
fn test_js_class_expression_async_methods() {
    let src = r#"
const AsyncClass = class {
    async fetch() { return "AsyncResult"; }
};
new AsyncClass().fetch().then(res => console.log(res));
"#;
    assert_eq!(run_js(src), vec!["AsyncResult"]);
}

#[test]
fn test_js_class_expression_symbol_iterator_protocol() {
    let src = r#"
const IterableClass = class {
    *[Symbol.iterator]() {
        yield "X"; yield "Y";
    }
};
console.log([...new IterableClass()].join("-"));
"#;
    assert_eq!(run_js(src), vec!["X-Y"]);
}

#[test]
fn test_js_class_expression_destructuring_assignment() {
    let src = r#"
const { Model } = {
    Model: class {
        constructor(id) { this.id = id; }
    }
};
console.log(new Model(42).id);
"#;
    assert_eq!(run_js(src), vec!["42"]);
}

#[test]
fn test_js_class_expression_typeof_operator() {
    let src = r#"
const Expr = class {};
console.log(typeof Expr);
"#;
    assert_eq!(run_js(src), vec!["function"]);
}
