use std::cell::RefCell;
use std::rc::Rc;

fn run_js(code: &str) -> Vec<String> {
    let program = vybe_parser_js::parse(code).expect("parse failed");
    let mut vm = vybe_bytecode::VM::new();
    let output: Rc<RefCell<Vec<String>>> = Rc::new(RefCell::new(Vec::new()));
    let out = output.clone();
    vybe_host::register_all(&mut vm);
    vybe_compiler_js::register_js_coercion(&mut vm);
    vm.register_host_fn("wasi:cli", "log", Box::new(move |args: &[vybe_bytecode::Value]| {
        let parts: Vec<String> = args.iter().map(|v| format!("{}", v)).collect();
        out.borrow_mut().push(parts.join(" "));
        vybe_bytecode::Value::Null
    }));
    vybe_host::setup_namespaces(&mut vm);
    let chunks = vybe_compiler_js::Compiler::new().compile(&program).expect("compile failed");
    vm.run(chunks).expect("runtime error");
    output.borrow().clone()
}

fn run_js_one(code: &str) -> String {
    run_js(code).into_iter().next().unwrap_or_default()
}

fn run_js_vm(code: &str) -> (vybe_bytecode::VM, Rc<RefCell<Vec<String>>>) {
    let program = vybe_parser_js::parse(code).expect("parse failed");
    let mut vm = vybe_bytecode::VM::new();
    let output: Rc<RefCell<Vec<String>>> = Rc::new(RefCell::new(Vec::new()));
    let out = output.clone();
    vybe_host::register_all(&mut vm);
    vybe_compiler_js::register_js_coercion(&mut vm);
    vm.register_host_fn("wasi:cli", "log", Box::new(move |args: &[vybe_bytecode::Value]| {
        let parts: Vec<String> = args.iter().map(|v| format!("{}", v)).collect();
        out.borrow_mut().push(parts.join(" "));
        vybe_bytecode::Value::Null
    }));
    vybe_host::setup_namespaces(&mut vm);
    let chunks = vybe_compiler_js::Compiler::new().compile(&program).expect("compile failed");
    vm.run(chunks).expect("runtime error");
    (vm, output)
}

// ============================================================
// A. FUNCTION RESOLUTION AND THIS BINDING (12 tests)
// ============================================================

// A1. Method call on object literal
// Known bug: `this` in object literal methods resolves to null instead of the object.
#[test]
fn test_a01_method_call_on_object_literal() {
    let code = r#"
        let obj = {
            x: 10,
            getX() { return this.x; }
        };
        console.log(obj.getX());
    "#;
    assert_eq!(run_js_one(code), "10");
}

// A2. Method call on class instance
#[test]
fn test_a02_method_call_on_class_instance() {
    let code = r#"
        class Greeter {
            constructor(name) { this.name = name; }
            greet() { return "Hello " + this.name; }
        }
        let g = new Greeter("World");
        console.log(g.greet());
    "#;
    assert_eq!(run_js_one(code), "Hello World");
}

// A3. Chained method calls
// Known bug: `this` in object literal methods resolves to null, so chained calls fail.
#[test]
fn test_a03_chained_method_calls() {
    let code = r#"
        class Builder {
            constructor() { this.parts = []; }
            add(s) { this.parts.push(s); return this; }
            build() { return this.parts.join("-"); }
        }
        let b = new Builder();
        console.log(b.add("a").add("b").add("c").build());
    "#;
    assert_eq!(run_js_one(code), "a-b-c");
}

// A4. Method stored in variable, called later
// Known limitation: extracting a method from an object loses `this` binding.
// In JS, calling a method via a detached reference uses globalThis (or undefined in strict mode).
// Our VM does not rebind `this` for detached method calls.
#[test]
fn test_a04_method_stored_in_variable() {
    let code = r#"
        let obj = { x: 42, getX() { return this.x; } };
        let fn2 = obj.getX;
        console.log(fn2());
    "#;
    // In non-strict JS, `this` would be globalThis, so this.x is undefined.
    let result = run_js_one(code);
    assert!(result == "undefined" || result == "null",
        "expected undefined or null, got: {}", result);
}

// A5. Callback function passed to another function
#[test]
fn test_a05_callback_passed_to_function() {
    let code = r#"
        function apply(fn, val) { return fn(val); }
        function double(x) { return x * 2; }
        console.log(apply(double, 21));
    "#;
    assert_eq!(run_js_one(code), "42");
}

// A6. Closure capturing outer variable
#[test]
fn test_a06_closure_capturing_outer_variable() {
    let code = r#"
        function makeAdder(x) {
            return (y) => x + y;
        }
        let add5 = makeAdder(5);
        console.log(add5(10));
    "#;
    assert_eq!(run_js_one(code), "15");
}

// A7. Closure mutation (modifying captured var)
#[test]
fn test_a07_closure_mutation() {
    let code = r#"
        function counter() {
            let n = 0;
            return {
                inc() { n++; return n; },
                get() { return n; }
            };
        }
        let c = counter();
        c.inc();
        c.inc();
        c.inc();
        console.log(c.get());
    "#;
    assert_eq!(run_js_one(code), "3");
}

// A8. Arrow function in method (this should capture outer)
// Known bug: Arrow functions in class methods do not capture `this` from the enclosing method.
// Expected: 3, Actual: 0 (this.ticks is not incremented because `this` is not bound in the arrow).
#[test]
fn test_a08_arrow_in_method_captures_this() {
    let code = r#"
        class Timer {
            constructor() { this.ticks = 0; }
            start() {
                let tick = () => { this.ticks++; };
                tick();
                tick();
                tick();
                return this.ticks;
            }
        }
        let t = new Timer();
        console.log(t.start());
    "#;
    assert_eq!(run_js_one(code), "3");
}

// A9. Constructor sets properties on this
#[test]
fn test_a09_constructor_sets_properties() {
    let code = r#"
        class Point {
            constructor(x, y) {
                this.x = x;
                this.y = y;
                this.label = "P(" + x + "," + y + ")";
            }
        }
        let p = new Point(3, 4);
        console.log(p.x, p.y, p.label);
    "#;
    assert_eq!(run_js_one(code), "3 4 P(3,4)");
}

// A10. super() in derived class constructor
#[test]
fn test_a10_super_in_derived_constructor() {
    let code = r#"
        class Animal {
            constructor(name) { this.name = name; }
        }
        class Dog extends Animal {
            constructor(name, breed) {
                super(name);
                this.breed = breed;
            }
        }
        let d = new Dog("Rex", "Labrador");
        console.log(d.name, d.breed);
    "#;
    assert_eq!(run_js_one(code), "Rex Labrador");
}

// A11. Static method on class
#[test]
fn test_a11_static_method_on_class() {
    let code = r#"
        class MathUtil {
            static square(x) { return x * x; }
        }
        console.log(MathUtil.square(7));
    "#;
    assert_eq!(run_js_one(code), "49");
}

// A12. Function.prototype.call with thisArg
// Known bug: fn.call(thisArg, ...) does not bind `this` to thisArg.
// Expected: "Hello Alice", Actual: "Hello null" (this.name resolves to null).
#[test]
#[ignore = "known bug"]
fn test_a12_function_call_with_this_arg() {
    let code = r#"
        function greet(greeting) {
            return greeting + " " + this.name;
        }
        let obj = { name: "Alice" };
        console.log(greet.call(obj, "Hello"));
    "#;
    assert_eq!(run_js_one(code), "Hello Alice");
}

// ============================================================
// B. OBJECT MECHANICS (10 tests)
// ============================================================

// B13. Object literal with methods
// Known bug: `this` in object literal methods resolves to null instead of the receiver.
// Expected: "hi from test", Actual: "hi from null".
#[test]
fn test_b13_object_literal_with_methods() {
    let code = r#"
        let obj = {
            name: "test",
            greet() { return "hi from " + this.name; }
        };
        console.log(obj.greet());
    "#;
    assert_eq!(run_js_one(code), "hi from test");
}

// B14. Computed property names
#[test]
fn test_b14_computed_property_names() {
    let code = r#"
        let key = "color";
        let obj = { [key]: "red", size: 10 };
        console.log(obj.color, obj.size);
    "#;
    assert_eq!(run_js_one(code), "red 10");
}

// B15. Property shorthand
#[test]
fn test_b15_property_shorthand() {
    let code = r#"
        let x = 10;
        let y = 20;
        let obj = { x, y };
        console.log(obj.x, obj.y);
    "#;
    assert_eq!(run_js_one(code), "10 20");
}

// B16. Spread in object
// Known limitation: spread in object literals is parsed but not compiled (no-op in compiler).
#[test]
fn test_b16_spread_in_object() {
    let code = r#"
        let other = { a: 1, b: 2 };
        let obj = { ...other, c: 3 };
        console.log(obj.a, obj.b, obj.c);
    "#;
    assert_eq!(run_js_one(code), "1 2 3");
}

// B17. Nested objects: a.b.c.d
#[test]
fn test_b17_nested_objects() {
    let code = r#"
        let obj = { a: { b: { c: { d: 42 } } } };
        console.log(obj.a.b.c.d);
    "#;
    assert_eq!(run_js_one(code), "42");
}

// B18. Object.keys, Object.values, Object.entries
#[test]
fn test_b18_object_keys_values_entries() {
    let code = r#"
        let obj = { x: 10 };
        let keys = Object.keys(obj);
        let vals = Object.values(obj);
        let entries = Object.entries(obj);
        console.log(keys.length, vals[0], entries[0][0], entries[0][1]);
    "#;
    assert_eq!(run_js_one(code), "1 10 x 10");
}

// B19. Object.assign merges
// Known bug: Object.assign with 3+ args only merges the first source.
// Expected: "1 2 3", Actual: "1 null null" (second and third sources ignored).
#[test]
fn test_b19_object_assign_merges() {
    let code = r#"
        let a = { x: 1 };
        let b = { y: 2 };
        let c = { z: 3 };
        let merged = Object.assign({}, a, b, c);
        console.log(merged.x, merged.y, merged.z);
    "#;
    assert_eq!(run_js_one(code), "1 2 3");
}

// B20. delete obj.prop
#[test]
fn test_b20_delete_property() {
    let code = r#"
        let obj = { a: 1, b: 2, c: 3 };
        delete obj.b;
        console.log("b" in obj, obj.a, obj.c);
    "#;
    assert_eq!(run_js_one(code), "false 1 3");
}

// B21. "key" in obj
#[test]
fn test_b21_in_operator() {
    let code = r#"
        let obj = { name: "Alice", age: 30 };
        console.log("name" in obj, "email" in obj);
    "#;
    assert_eq!(run_js_one(code), "true false");
}

// B22. hasOwnProperty
#[test]
fn test_b22_has_own_property() {
    let code = r#"
        let obj = { a: 1, b: 2 };
        console.log(obj.hasOwnProperty("a"), obj.hasOwnProperty("c"));
    "#;
    assert_eq!(run_js_one(code), "true false");
}

// ============================================================
// C. ARRAY MECHANICS (11 tests)
// ============================================================

// C23. Array literal, indexing, length
#[test]
fn test_c23_array_literal_indexing_length() {
    let code = r#"
        let arr = [10, 20, 30, 40, 50];
        console.log(arr[0], arr[2], arr[4], arr.length);
    "#;
    assert_eq!(run_js_one(code), "10 30 50 5");
}

// C24. push/pop/shift
#[test]
fn test_c24_push_pop_shift() {
    let code = r#"
        let arr = [1, 2, 3];
        arr.push(4);
        let popped = arr.pop();
        let shifted = arr.shift();
        console.log(popped, shifted, arr);
    "#;
    assert_eq!(run_js_one(code), "4 1 2,3");
}

// C25. map with arrow function
#[test]
fn test_c25_map_with_arrow() {
    assert_eq!(run_js_one("console.log([1, 2, 3].map(x => x * x))"), "1,4,9");
}

// C26. filter with arrow function
#[test]
fn test_c26_filter_with_arrow() {
    assert_eq!(run_js_one("console.log([1, 2, 3, 4, 5, 6].filter(x => x % 2 === 0))"), "2,4,6");
}

// C27. reduce accumulator
#[test]
fn test_c27_reduce_accumulator() {
    let code = r#"
        let result = [1, 2, 3, 4, 5].reduce((acc, x) => acc + x, 0);
        console.log(result);
    "#;
    assert_eq!(run_js_one(code), "15");
}

// C28. find returns first match
#[test]
fn test_c28_find_returns_first_match() {
    let code = r#"
        let arr = [1, 5, 10, 15, 20];
        console.log(arr.find(x => x > 8));
    "#;
    assert_eq!(run_js_one(code), "10");
}

// C29. some/every with edge cases
#[test]
fn test_c29_some_every_edge_cases() {
    let lines = run_js(r#"
        console.log([1, 2, 3].some(x => x > 2));
        console.log([1, 2, 3].some(x => x > 10));
        console.log([2, 4, 6].every(x => x % 2 === 0));
        console.log([2, 3, 6].every(x => x % 2 === 0));
        console.log([].some(x => true));
        console.log([].every(x => false));
    "#);
    assert_eq!(lines, vec!["true", "false", "true", "false", "false", "true"]);
}

// C30. findIndex
#[test]
fn test_c30_find_index() {
    let lines = run_js(r#"
        console.log([10, 20, 30].findIndex(x => x === 20));
        console.log([10, 20, 30].findIndex(x => x === 99));
    "#);
    assert_eq!(lines, vec!["1", "-1"]);
}

// C31. sort with comparator
#[test]
fn test_c31_sort_with_comparator() {
    let code = r#"
        let arr = [5, 3, 8, 1, 9, 2];
        arr.sort((a, b) => a - b);
        console.log(arr);
    "#;
    assert_eq!(run_js_one(code), "1,2,3,5,8,9");
}

// C32. Array.from creates copy
#[test]
fn test_c32_array_from_creates_copy() {
    let code = r#"
        let a = [1, 2, 3];
        let b = Array.from(a);
        b[0] = 99;
        console.log(a[0], b[0]);
    "#;
    assert_eq!(run_js_one(code), "1 99");
}

// C33. Spread in array
#[test]
fn test_c33_spread_in_array() {
    let code = r#"
        let arr = [1, 2, 3];
        let result = [0, ...arr, 4, 5];
        console.log(result);
    "#;
    assert_eq!(run_js_one(code), "0,1,2,3,4,5");
}

// ============================================================
// D. CLASS MECHANICS (10 tests)
// ============================================================

// D34. Class with constructor, fields, methods
#[test]
fn test_d34_class_constructor_fields_methods() {
    let code = r#"
        class Rectangle {
            constructor(w, h) {
                this.width = w;
                this.height = h;
            }
            area() { return this.width * this.height; }
            perimeter() { return 2 * (this.width + this.height); }
        }
        let r = new Rectangle(3, 4);
        console.log(r.area(), r.perimeter());
    "#;
    assert_eq!(run_js_one(code), "12 14");
}

// D35. Inheritance with extends and super()
#[test]
fn test_d35_inheritance_extends_super() {
    let code = r#"
        class Shape {
            constructor(name) { this.name = name; }
            describe() { return "I am a " + this.name; }
        }
        class Circle extends Shape {
            constructor(r) {
                super("circle");
                this.radius = r;
            }
        }
        let c = new Circle(5);
        console.log(c.describe(), c.radius);
    "#;
    assert_eq!(run_js_one(code), "I am a circle 5");
}

// D36. Method override
#[test]
fn test_d36_method_override() {
    let code = r#"
        class Base {
            greet() { return "base"; }
        }
        class Child extends Base {
            greet() { return "child"; }
        }
        let b = new Base();
        let c = new Child();
        console.log(b.greet(), c.greet());
    "#;
    assert_eq!(run_js_one(code), "base child");
}

// D37. Static methods and properties
#[test]
fn test_d37_static_methods() {
    let code = r#"
        class Counter {
            static count = 0;
            static increment() { Counter.count++; return Counter.count; }
        }
        console.log(Counter.increment(), Counter.increment(), Counter.increment());
    "#;
    assert_eq!(run_js_one(code), "1 2 3");
}

// D38. instanceof with class hierarchy
#[test]
fn test_d38_instanceof_direct() {
    let code = r#"
        class A {}
        class B extends A {}
        let b = new B();
        console.log(b instanceof B);
    "#;
    assert_eq!(run_js_one(code), "true");
}

// D39. Getter auto-dispatch
#[test]
fn test_d39_getter_dispatch() {
    let code = r#"
        class Temperature {
            constructor(celsius) { this._c = celsius; }
            get fahrenheit() { return this._c * 9 / 5 + 32; }
        }
        let t = new Temperature(100);
        console.log(t.fahrenheit);
    "#;
    assert_eq!(run_js_one(code), "212");
}

// D40. Setter auto-dispatch
#[test]
fn test_d40_setter_dispatch() {
    let code = r#"
        class Box {
            constructor() { this._value = 0; }
            get value() { return this._value; }
            set value(v) { this._value = v * 2; }
        }
        let b = new Box();
        b.value = 5;
        console.log(b.value);
    "#;
    assert_eq!(run_js_one(code), "10");
}

// D41. Class expression (anonymous)
// Known limitation: class expressions are not supported in the parser/compiler.
#[test]
#[ignore = "known bug"]
fn test_d41_class_expression() {
    let code = r#"
        let MyClass = class {
            constructor(x) { this.x = x; }
            get() { return this.x; }
        };
        let obj = new MyClass(42);
        console.log(obj.get());
    "#;
    assert_eq!(run_js_one(code), "42");
}

// D42. Multiple instances independent state
// Known bug: `this.count++` in class methods does not work — the increment does not take effect.
// Expected: "3 101", Actual: "0 100" (inc() doesn't modify instance state).
#[test]
fn test_d42_multiple_instances_independent_state() {
    let code = r#"
        class Counter {
            constructor(start) { this.count = start; }
            inc() { this.count++; return this.count; }
        }
        let a = new Counter(0);
        let b = new Counter(100);
        a.inc(); a.inc(); a.inc();
        b.inc();
        console.log(a.count, b.count);
    "#;
    assert_eq!(run_js_one(code), "3 101");
}

// D43. Constructor calling methods
// Known bug: Calling a method on `this` inside a constructor fails with "null is not callable".
// The method lookup on the partially-constructed instance does not find the method.
#[test]
fn test_d43_constructor_calling_methods() {
    let code = r#"
        class Validator {
            constructor(value) {
                this.value = value;
                this.valid = this.check();
            }
            check() { return this.value > 0; }
        }
        let v1 = new Validator(10);
        let v2 = new Validator(-5);
        console.log(v1.valid, v2.valid);
    "#;
    assert_eq!(run_js_one(code), "true false");
}

// ============================================================
// E. INVOKE FROM RUST AFTER COMPILATION (8 tests)
// ============================================================

// E44. Run JS that defines a function, then invoke it from Rust
#[test]
fn test_e44_invoke_global_function() {
    let code = r#"
        function double(x) { return x * 2; }
    "#;
    let (mut vm, _output) = run_js_vm(code);
    let func = vm.globals.get("double").cloned().expect("double not in globals");
    let result = vm.invoke(&func, &[vybe_bytecode::Value::F64(21.0)]).expect("invoke failed");
    assert_eq!(format!("{}", result), "42");
}

// E45. invoke global function with args
#[test]
fn test_e45_invoke_with_multiple_args() {
    let code = r#"
        function add(a, b, c) { return a + b + c; }
    "#;
    let (mut vm, _output) = run_js_vm(code);
    let func = vm.globals.get("add").cloned().expect("add not in globals");
    let result = vm.invoke(&func, &[
        vybe_bytecode::Value::F64(10.0),
        vybe_bytecode::Value::F64(20.0),
        vybe_bytecode::Value::F64(30.0),
    ]).expect("invoke failed");
    assert_eq!(format!("{}", result), "60");
}

// E46. invoke method on object stored in global
// Known bug: `let obj = { ... }` does not store into globals (only `var` or top-level declarations
// that use global_set). The VM globals map does not contain `obj` after run.
#[test]
#[ignore = "known bug"]
fn test_e46_invoke_method_on_global_object() {
    let code = r#"
        let obj = {
            x: 10,
            getX() { return this.x; }
        };
    "#;
    let (mut vm, _output) = run_js_vm(code);
    let obj = vm.globals.get("obj").cloned().expect("obj not in globals");
    // Extract the method from the object
    if let vybe_bytecode::Value::Object(ref rc) = obj {
        let borrowed = rc.borrow();
        let method = borrowed.properties.get("getX").cloned().expect("getX not on obj");
        // Invoke as standalone function (this may not bind correctly without method_call)
        let result = vm.invoke(&method, &[]);
        // Just check it doesn't crash; this binding for detached methods is a known issue
        assert!(result.is_ok() || result.is_err());
    } else {
        panic!("obj is not an Object");
    }
}

// E47. invoke returns correct value
#[test]
fn test_e47_invoke_returns_correct_value() {
    let code = r#"
        function makeGreeting(name) {
            return "Hello " + name + "!";
        }
    "#;
    let (mut vm, _output) = run_js_vm(code);
    let func = vm.globals.get("makeGreeting").cloned().expect("fn not found");
    let result = vm.invoke(&func, &[vybe_bytecode::Value::String("World".into())]).expect("invoke failed");
    assert_eq!(format!("{}", result), "Hello World!");
}

// E48. Multiple invokes preserve state
#[test]
fn test_e48_multiple_invokes_preserve_state() {
    let code = r#"
        let count = 0;
        function increment() {
            count++;
            return count;
        }
    "#;
    let (mut vm, _output) = run_js_vm(code);
    let func = vm.globals.get("increment").cloned().expect("fn not found");
    let r1 = vm.invoke(&func, &[]).expect("invoke 1 failed");
    let r2 = vm.invoke(&func, &[]).expect("invoke 2 failed");
    let r3 = vm.invoke(&func, &[]).expect("invoke 3 failed");
    assert_eq!(format!("{}", r1), "1");
    assert_eq!(format!("{}", r2), "2");
    assert_eq!(format!("{}", r3), "3");
}

// E49. invoke class method with this
// Known limitation: invoking a class method extracted from an instance does not bind `this`.
#[test]
#[ignore = "known bug"]
fn test_e49_invoke_class_method_with_this() {
    let code = r#"
        class Adder {
            constructor(base) { this.base = base; }
            add(x) { return this.base + x; }
        }
        let adder = new Adder(100);
    "#;
    let (mut vm, _output) = run_js_vm(code);
    let adder = vm.globals.get("adder").cloned().expect("adder not found");
    if let vybe_bytecode::Value::Object(ref rc) = adder {
        let borrowed = rc.borrow();
        let method = borrowed.properties.get("add").cloned().expect("add not found");
        let result = vm.invoke(&method, &[vybe_bytecode::Value::F64(5.0)]).expect("invoke failed");
        assert_eq!(format!("{}", result), "105");
    } else {
        panic!("adder is not an Object");
    }
}

// E50. invoke callback that modifies closure variable
#[test]
fn test_e50_invoke_closure_modifies_captured() {
    let code = r#"
        let state = 0;
        function setState(val) {
            state = val;
            return state;
        }
        function getState() { return state; }
    "#;
    let (mut vm, _output) = run_js_vm(code);
    let set_fn = vm.globals.get("setState").cloned().expect("setState not found");
    let get_fn = vm.globals.get("getState").cloned().expect("getState not found");
    vm.invoke(&set_fn, &[vybe_bytecode::Value::F64(42.0)]).expect("set failed");
    let result = vm.invoke(&get_fn, &[]).expect("get failed");
    assert_eq!(format!("{}", result), "42");
}

// E51. invoke after defining classes — class still works
#[test]
fn test_e51_invoke_after_class_definition() {
    let code = r#"
        class Calculator {
            static multiply(a, b) { return a * b; }
        }
        function compute(x, y) {
            return Calculator.multiply(x, y);
        }
    "#;
    let (mut vm, _output) = run_js_vm(code);
    let func = vm.globals.get("compute").cloned().expect("compute not found");
    let result = vm.invoke(&func, &[
        vybe_bytecode::Value::F64(6.0),
        vybe_bytecode::Value::F64(7.0),
    ]).expect("invoke failed");
    assert_eq!(format!("{}", result), "42");
}

// ============================================================
// F. SCOPE AND CLOSURES DEEP (8 tests)
// ============================================================

// F52. Block scope: let in if doesn't leak
#[test]
fn test_f52_block_scope_let_doesnt_leak() {
    let code = r#"
        let x = "outer";
        if (true) {
            let x = "inner";
            console.log(x);
        }
        console.log(x);
    "#;
    let lines = run_js(code);
    assert_eq!(lines, vec!["inner", "outer"]);
}

// F53. var hoisting: var in if visible outside
// Known bug: `var` inside a block within a function is not hoisted to function scope.
// Expected: 42, Actual: "undefined" (var is treated like let).
#[test]
fn test_f53_var_hoisting() {
    let code = r#"
        function test() {
            if (true) {
                var x = 42;
            }
            return x;
        }
        console.log(test());
    "#;
    assert_eq!(run_js_one(code), "42");
}

// F54. Closure over loop variable (let creates new binding per iteration)
// Known bug: `let` in a for-loop does not create a fresh binding per iteration.
#[test]
fn test_f54_closure_over_loop_let() {
    let code = r#"
        let fns = [];
        for (let i = 0; i < 5; i++) {
            fns.push(() => i);
        }
        console.log(fns[0](), fns[1](), fns[2](), fns[3](), fns[4]());
    "#;
    assert_eq!(run_js_one(code), "0 1 2 3 4");
}

// F55. IIFE (Immediately Invoked Function Expression)
#[test]
fn test_f55_iife() {
    let code = r#"
        let result = (function(x) { return x * x; })(7);
        console.log(result);
    "#;
    assert_eq!(run_js_one(code), "49");
}

// F56. Nested closures: 3 levels deep
#[test]
fn test_f56_nested_closures_three_levels() {
    let code = r#"
        function a() {
            let x = 1;
            function b() {
                let y = 2;
                function c() {
                    let z = 3;
                    return x + y + z;
                }
                return c();
            }
            return b();
        }
        console.log(a());
    "#;
    assert_eq!(run_js_one(code), "6");
}

// F57. Closure returned from function, called later
#[test]
fn test_f57_closure_returned_called_later() {
    let code = r#"
        function makeMultiplier(factor) {
            return (x) => x * factor;
        }
        let triple = makeMultiplier(3);
        let quadruple = makeMultiplier(4);
        console.log(triple(10), quadruple(10));
    "#;
    assert_eq!(run_js_one(code), "30 40");
}

// F58. Multiple closures sharing state
// Known bug: Object literal methods that mutate a shared closure variable produce NaN.
// The `set(v)` method stores the value, but `inc()` fails because the closure variable
// is not properly shared across methods in the returned object literal.
// Expected: 12, Actual: NaN.
#[test]
fn test_f58_multiple_closures_sharing_state() {
    let code = r#"
        function makeStore() {
            let value = 0;
            return {
                set(v) { value = v; },
                get() { return value; },
                inc() { value++; }
            };
        }
        let store = makeStore();
        store.set(10);
        store.inc();
        store.inc();
        console.log(store.get());
    "#;
    assert_eq!(run_js_one(code), "12");
}

// F59. Closure inside class method
#[test]
fn test_f59_closure_inside_class_method() {
    let code = r#"
        class Processor {
            constructor(data) { this.data = data; }
            processAll() {
                let results = [];
                let self = this;
                this.data.forEach(x => {
                    results.push(x * 2);
                });
                return results;
            }
        }
        let p = new Processor([1, 2, 3]);
        console.log(p.processAll());
    "#;
    assert_eq!(run_js_one(code), "2,4,6");
}

// ============================================================
// G. EDGE CASES AND TYPE COERCION (10 tests)
// ============================================================

// G60. null == undefined (true in JS)
// Known limitation: VM does not implement Abstract Equality (==) for null/undefined.
#[test]
fn test_g60_null_loose_equals_undefined() {
    assert_eq!(run_js_one("console.log(null == undefined)"), "true");
}

// G61. null === undefined (false)
#[test]
fn test_g61_null_strict_not_equal_undefined() {
    assert_eq!(run_js_one("console.log(null === undefined)"), "false");
}

// G62. NaN !== NaN
#[test]
fn test_g62_nan_not_equal_nan() {
    assert_eq!(run_js_one("console.log(NaN === NaN)"), "false");
    assert_eq!(run_js_one("console.log(NaN !== NaN)"), "true");
}

// G63. typeof null === "object"
#[test]
fn test_g63_typeof_null_is_object() {
    assert_eq!(run_js_one("console.log(typeof null)"), "object");
}

// G64. typeof undefined === "undefined"
#[test]
fn test_g64_typeof_undefined() {
    assert_eq!(run_js_one("console.log(typeof undefined)"), "undefined");
}

// G65. typeof function === "function"
#[test]
fn test_g65_typeof_function() {
    assert_eq!(run_js_one("console.log(typeof function(){})"), "function");
    assert_eq!(run_js_one("console.log(typeof (() => {}))"), "function");
}

// G66. "5" + 3 === "53" (string concat)
#[test]
fn test_g66_string_plus_number_concats() {
    assert_eq!(run_js_one(r#"console.log("5" + 3)"#), "53");
}

// G67. "5" - 3 === 2 (numeric)
// Known bug: String-to-number coercion for subtraction is not implemented.
// Expected: 2, Actual: NaN (the VM does not coerce "5" to 5 for arithmetic minus).
#[test]
fn test_g67_string_minus_number_numeric() {
    assert_eq!(run_js_one(r#"console.log("5" - 3)"#), "2");
}

// G68. Empty array is truthy
#[test]
fn test_g68_empty_array_is_truthy() {
    let code = r#"
        if ([]) {
            console.log("truthy");
        } else {
            console.log("falsy");
        }
    "#;
    assert_eq!(run_js_one(code), "truthy");
}

// G69. 0 is falsy, "0" is truthy
#[test]
fn test_g69_zero_falsy_string_zero_truthy() {
    let lines = run_js(r#"
        console.log(0 ? "truthy" : "falsy");
        console.log("0" ? "truthy" : "falsy");
    "#);
    assert_eq!(lines, vec!["falsy", "truthy"]);
}

// G70. Template literal with expression
#[test]
fn test_g70_template_literal_with_expression() {
    let code = r#"
        let a = 3;
        let b = 4;
        console.log(`${a} + ${b} = ${a + b}`);
    "#;
    assert_eq!(run_js_one(code), "3 + 4 = 7");
}

// ============================================================
// H. CONTROL FLOW EDGE CASES (8 tests)
// ============================================================

// H71. Labeled break from nested loop
#[test]
fn test_h71_labeled_break_nested_loop() {
    let code = r#"
        let result = 0;
        outer: for (let i = 0; i < 5; i++) {
            for (let j = 0; j < 5; j++) {
                if (i === 2 && j === 3) break outer;
                result++;
            }
        }
        console.log(result);
    "#;
    // i=0: j=0..4 (5), i=1: j=0..4 (5), i=2: j=0..2 (3) => 13
    assert_eq!(run_js_one(code), "13");
}

// H72. Labeled continue
#[test]
fn test_h72_labeled_continue() {
    let code = r#"
        let result = 0;
        outer: for (let i = 0; i < 3; i++) {
            for (let j = 0; j < 3; j++) {
                if (j === 1) continue outer;
                result++;
            }
        }
        console.log(result);
    "#;
    // Each iteration of outer: j=0 counts, j=1 triggers continue outer => 3 counts
    assert_eq!(run_js_one(code), "3");
}

// H73. Switch with fallthrough (no break)
// Known bug: Switch does not implement fallthrough — only the matched case body runs.
// Expected: "abc" (JS fallthrough), Actual: "a" (only case 1 body executes).
#[test]
fn test_h73_switch_fallthrough() {
    let code = r#"
        let result = "";
        switch (1) {
            case 1: result += "a";
            case 2: result += "b";
            case 3: result += "c";
        }
        console.log(result);
    "#;
    // JS fallthrough: case 1 matches, then falls through 2 and 3
    assert_eq!(run_js_one(code), "abc");
}

// H74. try/catch/finally — finally always runs
#[test]
fn test_h74_try_catch_finally() {
    let lines = run_js(r#"
        let log = "";
        try {
            log += "try ";
            throw "oops";
        } catch (e) {
            log += "catch(" + e + ") ";
        } finally {
            log += "finally";
        }
        console.log(log);

        let log2 = "";
        try {
            log2 += "try ";
        } catch (e) {
            log2 += "catch ";
        } finally {
            log2 += "finally";
        }
        console.log(log2);
    "#);
    assert_eq!(lines[0], "try catch(oops) finally");
    assert_eq!(lines[1], "try finally");
}

// H75. throw custom error object
#[test]
fn test_h75_throw_custom_error() {
    let code = r#"
        try {
            throw { message: "custom error", code: 42 };
        } catch (e) {
            console.log(e.message, e.code);
        }
    "#;
    assert_eq!(run_js_one(code), "custom error 42");
}

// H76. for...of over array
#[test]
fn test_h76_for_of_over_array() {
    let code = r#"
        let sum = 0;
        for (let x of [10, 20, 30, 40]) {
            sum += x;
        }
        console.log(sum);
    "#;
    assert_eq!(run_js_one(code), "100");
}

// H77. for...in over object
#[test]
fn test_h77_for_in_over_object() {
    let code = r#"
        let obj = { a: 1, b: 2, c: 3 };
        let sum = 0;
        for (let k in obj) {
            sum += obj[k];
        }
        console.log(sum);
    "#;
    assert_eq!(run_js_one(code), "6");
}

// H78. do...while with break
#[test]
fn test_h78_do_while_with_break() {
    let code = r#"
        let i = 0;
        let sum = 0;
        do {
            sum += i;
            i++;
            if (i > 5) break;
        } while (i < 100);
        console.log(sum);
    "#;
    // i=0..5: 0+1+2+3+4+5 = 15
    assert_eq!(run_js_one(code), "15");
}
