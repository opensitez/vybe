use super::helpers::run_js;

// ===================================================================
// 1. Private field basic access via method
// ===================================================================

#[test]
fn private_field_basic_access_via_method() {
    assert_eq!(
        run_js(
            r#"
class BankAccount {
    #balance = 100;
    getBalance() { return this.#balance; }
}
const acc = new BankAccount();
console.log(acc.getBalance());
"#
        ),
        vec!["100"]
    );
}

// ===================================================================
// 2. Private method accessible within class only
// ===================================================================

#[test]
fn private_method_accessible_within_class_only() {
    assert_eq!(
        run_js(
            r#"
class Validator {
    #isNonEmpty(s) { return s.length > 0; }
    validate(s) { return this.#isNonEmpty(s) ? "ok" : "empty"; }
}
const v = new Validator();
console.log(v.validate("hello"));
console.log(v.validate(""));
"#
        ),
        vec!["ok", "empty"]
    );
}

// ===================================================================
// 3. Private static field shared across instances
// ===================================================================

#[test]
fn private_static_field_shared_across_instances() {
    assert_eq!(
        run_js(
            r#"
class Registry {
    static #count = 0;
    constructor() { Registry.#count++; }
    static getCount() { return Registry.#count; }
}
new Registry();
new Registry();
new Registry();
console.log(Registry.getCount());
"#
        ),
        vec!["3"]
    );
}

// ===================================================================
// 4. Private static method callable on class
// ===================================================================

#[test]
fn private_static_method_callable_on_class() {
    assert_eq!(
        run_js(
            r#"
class MathHelper {
    static #double(x) { return x * 2; }
    static compute(x) { return MathHelper.#double(x) + 1; }
}
console.log(MathHelper.compute(5));
console.log(MathHelper.compute(10));
"#
        ),
        vec!["11", "21"]
    );
}

// ===================================================================
// 5. Static initialization block runs once at class definition
// ===================================================================

#[test]
fn static_init_block_runs_once_at_class_definition() {
    assert_eq!(
        run_js(
            r#"
let ran = 0;
class Setup {
    static {
        ran++;
    }
}
console.log(ran);
const a = new Setup();
const b = new Setup();
console.log(ran);
"#
        ),
        vec!["1", "1"]
    );
}

// ===================================================================
// 6. Static initialization block can call static methods
// ===================================================================

#[test]
fn static_init_block_can_call_static_methods() {
    assert_eq!(
        run_js(
            r#"
class Config {
    static value;
    static #compute() { return 42; }
    static {
        Config.value = Config.#compute();
    }
}
console.log(Config.value);
"#
        ),
        vec!["42"]
    );
}

// ===================================================================
// 7. Static initialization block can initialize static fields
// ===================================================================

#[test]
fn static_init_block_can_initialize_static_fields() {
    assert_eq!(
        run_js(
            r#"
class Constants {
    static PI;
    static E;
    static {
        Constants.PI = 3.14159;
        Constants.E = 2.71828;
    }
}
console.log(Constants.PI);
console.log(Constants.E);
"#
        ),
        vec!["3.14159", "2.71828"]
    );
}

// ===================================================================
// 8. Static initialization block order (multiple blocks run top to bottom)
// ===================================================================

#[test]
fn static_init_block_order_multiple_blocks() {
    assert_eq!(
        run_js(
            r#"
class Ordered {
    static log = [];
    static { Ordered.log.push("first"); }
    static { Ordered.log.push("second"); }
    static { Ordered.log.push("third"); }
}
console.log(Ordered.log.join(","));
"#
        ),
        vec!["first,second,third"]
    );
}

// ===================================================================
// 9. Private field in subclass is separate from parent's private field
// ===================================================================

#[test]
fn private_field_in_subclass_separate_from_parent() {
    assert_eq!(
        run_js(
            r#"
class Parent {
    #secret = "parent-secret";
    getSecret() { return this.#secret; }
}
class Child extends Parent {
    #secret = "child-secret";
    getChildSecret() { return this.#secret; }
}
const c = new Child();
console.log(c.getSecret());
console.log(c.getChildSecret());
"#
        ),
        vec!["parent-secret", "child-secret"]
    );
}

// ===================================================================
// 10. Private field existence check via try/catch
// ===================================================================

#[test]
fn private_field_access_outside_class_throws() {
    assert_eq!(
        run_js(
            r#"
class Secret {
    #value = 99;
    getValue() { return this.#value; }
}
const s = new Secret();
console.log(s.getValue());
try {
    console.log(s.#value);
} catch (e) {
    console.log("access denied");
}
"#
        ),
        vec!["99", "access denied"]
    );
}

// ===================================================================
// 11. Private getter property
// ===================================================================

#[test]
fn private_getter_property() {
    assert_eq!(
        run_js(
            r#"
class Circle {
    #radius;
    constructor(r) { this.#radius = r; }
    get #area() { return 3.14 * this.#radius * this.#radius; }
    describe() { return "area=" + this.#area; }
}
const c = new Circle(5);
console.log(c.describe());
"#
        ),
        vec!["area=78.5"]
    );
}

// ===================================================================
// 12. Private setter property
// ===================================================================

#[test]
fn private_setter_property() {
    assert_eq!(
        run_js(
            r#"
class Validated {
    #score = 0;
    set #safeScore(v) { this.#score = v < 0 ? 0 : v > 100 ? 100 : v; }
    setScore(v) { this.#safeScore = v; }
    getScore() { return this.#score; }
}
const obj = new Validated();
obj.setScore(150);
console.log(obj.getScore());
obj.setScore(-10);
console.log(obj.getScore());
obj.setScore(75);
console.log(obj.getScore());
"#
        ),
        vec!["100", "0", "75"]
    );
}

// ===================================================================
// 13. Private getter and setter pair
// ===================================================================

#[test]
fn private_getter_and_setter_pair() {
    assert_eq!(
        run_js(
            r#"
class Temperature {
    #celsius = 0;
    get #fahrenheit() { return this.#celsius * 9 / 5 + 32; }
    set #fahrenheit(f) { this.#celsius = (f - 32) * 5 / 9; }
    setF(f) { this.#fahrenheit = f; }
    getC() { return this.#celsius; }
    getF() { return this.#fahrenheit; }
}
const t = new Temperature();
t.setF(212);
console.log(t.getC());
console.log(t.getF());
"#
        ),
        vec!["100", "212"]
    );
}

// ===================================================================
// 14. Private field in toString method
// ===================================================================

#[test]
fn private_field_in_tostring_method() {
    assert_eq!(
        run_js(
            r#"
class Point {
    #x;
    #y;
    constructor(x, y) { this.#x = x; this.#y = y; }
    toString() { return "(" + this.#x + "," + this.#y + ")"; }
}
const p = new Point(3, 7);
console.log(p.toString());
console.log("Point: " + p);
"#
        ),
        vec!["(3,7)", "Point: (3,7)"]
    );
}

// ===================================================================
// 15. Private field in static factory method
// ===================================================================

#[test]
fn private_field_in_static_factory_method() {
    assert_eq!(
        run_js(
            r##"
class Color {
    #r; #g; #b;
    constructor(r, g, b) { this.#r = r; this.#g = g; this.#b = b; }
    static fromHex(hex) {
        const r = parseInt(hex.slice(1, 3), 16);
        const g = parseInt(hex.slice(3, 5), 16);
        const b = parseInt(hex.slice(5, 7), 16);
        return new Color(r, g, b);
    }
    toString() { return this.#r + "," + this.#g + "," + this.#b; }
}
const c = Color.fromHex("#ff8000");
console.log(c.toString());
"##
        ),
        vec!["255,128,0"]
    );
}

// ===================================================================
// 16. Class field (non-private public field)
// ===================================================================

#[test]
fn class_field_non_private_public() {
    assert_eq!(
        run_js(
            r#"
class Dog {
    species = "canine";
    constructor(name) { this.name = name; }
    describe() { return this.name + " is a " + this.species; }
}
const d = new Dog("Rex");
console.log(d.describe());
console.log(d.species);
"#
        ),
        vec!["Rex is a canine", "canine"]
    );
}

// ===================================================================
// 17. Public class field with default value
// ===================================================================

#[test]
fn public_class_field_with_default_value() {
    assert_eq!(
        run_js(
            r#"
class Task {
    done = false;
    priority = 1;
    label = "untitled";
}
const t = new Task();
console.log(t.done);
console.log(t.priority);
console.log(t.label);
"#
        ),
        vec!["false", "1", "untitled"]
    );
}

// ===================================================================
// 18. Public class field default overridden in constructor
// ===================================================================

#[test]
fn public_class_field_default_overridden_in_constructor() {
    assert_eq!(
        run_js(
            r#"
class Widget {
    color = "grey";
    constructor(color) {
        if (color) this.color = color;
    }
}
const w1 = new Widget("blue");
const w2 = new Widget();
console.log(w1.color);
console.log(w2.color);
"#
        ),
        vec!["blue", "grey"]
    );
}

// ===================================================================
// 19. Static public class field
// ===================================================================

#[test]
fn static_public_class_field() {
    assert_eq!(
        run_js(
            r#"
class App {
    static name = "MyApp";
    static version = "2.0";
    static description() { return App.name + " v" + App.version; }
}
console.log(App.name);
console.log(App.version);
console.log(App.description());
"#
        ),
        vec!["MyApp", "2.0", "MyApp v2.0"]
    );
}

// ===================================================================
// 20. Private field counter pattern (increment/decrement)
// ===================================================================

#[test]
fn private_field_counter_pattern() {
    assert_eq!(
        run_js(
            r#"
class Counter {
    #n = 0;
    inc() { this.#n++; return this; }
    dec() { this.#n--; return this; }
    reset() { this.#n = 0; return this; }
    value() { return this.#n; }
}
const c = new Counter();
c.inc().inc().inc().dec();
console.log(c.value());
c.reset();
console.log(c.value());
"#
        ),
        vec!["2", "0"]
    );
}

// ===================================================================
// 21. Private field for encapsulating state
// ===================================================================

#[test]
fn private_field_encapsulating_state() {
    assert_eq!(
        run_js(
            r#"
class TrafficLight {
    #state = "red";
    next() {
        if (this.#state === "red") this.#state = "green";
        else if (this.#state === "green") this.#state = "yellow";
        else this.#state = "red";
    }
    current() { return this.#state; }
}
const light = new TrafficLight();
console.log(light.current());
light.next();
console.log(light.current());
light.next();
console.log(light.current());
light.next();
console.log(light.current());
"#
        ),
        vec!["red", "green", "yellow", "red"]
    );
}

// ===================================================================
// 22. Class with both private and public fields accessing each other
// ===================================================================

#[test]
fn class_private_and_public_fields_access_each_other() {
    assert_eq!(
        run_js(
            r#"
class User {
    name;
    #role = "user";
    constructor(name, role) {
        this.name = name;
        if (role) this.#role = role;
    }
    display() { return this.name + " [" + this.#role + "]"; }
}
const u1 = new User("Alice", "admin");
const u2 = new User("Bob");
console.log(u1.display());
console.log(u2.display());
console.log(u1.name);
"#
        ),
        vec!["Alice [admin]", "Bob [user]", "Alice"]
    );
}

// ===================================================================
// 23. Private array field, public method to push/get
// ===================================================================

#[test]
fn private_array_field_public_push_get() {
    assert_eq!(
        run_js(
            r#"
class Stack {
    #items = [];
    push(item) { this.#items.push(item); }
    pop() { return this.#items.pop(); }
    size() { return this.#items.length; }
    peek() { return this.#items[this.#items.length - 1]; }
}
const s = new Stack();
s.push(10);
s.push(20);
s.push(30);
console.log(s.size());
console.log(s.peek());
console.log(s.pop());
console.log(s.size());
"#
        ),
        vec!["3", "30", "30", "2"]
    );
}

// ===================================================================
// 24. Private method calling another private method
// ===================================================================

#[test]
fn private_method_calling_another_private_method() {
    assert_eq!(
        run_js(
            r#"
class StringProcessor {
    #trim(s) { return s.trim(); }
    #upper(s) { return s.toUpperCase(); }
    #process(s) { return this.#upper(this.#trim(s)); }
    run(s) { return this.#process(s); }
}
const sp = new StringProcessor();
console.log(sp.run("  hello world  "));
"#
        ),
        vec!["HELLO WORLD"]
    );
}

// ===================================================================
// 25. Static private field for singleton count
// ===================================================================

#[test]
fn static_private_field_singleton_count() {
    assert_eq!(
        run_js(
            r#"
class Connection {
    static #pool = [];
    static #maxSize = 3;
    id;
    constructor(id) { this.id = id; }
    static acquire(id) {
        if (Connection.#pool.length < Connection.#maxSize) {
            const conn = new Connection(id);
            Connection.#pool.push(conn);
            return conn;
        }
        return null;
    }
    static poolSize() { return Connection.#pool.length; }
}
Connection.acquire("a");
Connection.acquire("b");
Connection.acquire("c");
const d = Connection.acquire("d");
console.log(Connection.poolSize());
console.log(d);
"#
        ),
        vec!["3", "null"]
    );
}

// ===================================================================
// 26. Private field serialization pattern (toJSON uses private fields)
// ===================================================================

#[test]
fn private_field_serialization_tojson() {
    assert_eq!(
        run_js(
            r#"
class Person {
    #name;
    #age;
    constructor(name, age) { this.#name = name; this.#age = age; }
    toJSON() { return { name: this.#name, age: this.#age }; }
    serialize() { return JSON.stringify(this.toJSON()); }
}
const p = new Person("Alice", 30);
const json = p.serialize();
console.log(json);
"#
        ),
        vec![r#"{"name":"Alice","age":30}"#]
    );
}

// ===================================================================
// 27. Class field declarations don't need constructor
// ===================================================================

#[test]
fn class_field_declarations_no_constructor_needed() {
    assert_eq!(
        run_js(
            r#"
class Config {
    host = "localhost";
    port = 8080;
    secure = false;
    toString() {
        const proto = this.secure ? "https" : "http";
        return proto + "://" + this.host + ":" + this.port;
    }
}
const cfg = new Config();
console.log(cfg.toString());
cfg.port = 443;
cfg.secure = true;
console.log(cfg.toString());
"#
        ),
        vec!["http://localhost:8080", "https://localhost:443"]
    );
}

// ===================================================================
// 28. instanceof with class having private fields
// ===================================================================

#[test]
fn instanceof_with_private_fields() {
    assert_eq!(
        run_js(
            r#"
class Animal {
    #alive = true;
    isAlive() { return this.#alive; }
}
class Dog extends Animal {
    #breed;
    constructor(breed) { super(); this.#breed = breed; }
    getBreed() { return this.#breed; }
}
const d = new Dog("Labrador");
console.log(d instanceof Dog);
console.log(d instanceof Animal);
console.log(d.isAlive());
console.log(d.getBreed());
"#
        ),
        vec!["true", "true", "true", "Labrador"]
    );
}

// ===================================================================
// 29. Private static method as helper in public static factory
// ===================================================================

#[test]
fn private_static_method_helper_in_public_factory() {
    assert_eq!(
        run_js(
            r#"
class UUID {
    static #pad(n) { return n.toString(16).padStart(4, "0"); }
    static #segment(max) { return Math.floor(max / 2); }
    static create(seed) {
        const a = UUID.#pad(seed);
        const b = UUID.#pad(UUID.#segment(seed));
        return a + "-" + b;
    }
}
console.log(UUID.create(256));
console.log(UUID.create(65536));
"#
        ),
        vec!["0100-0080", "10000-8000"]
    );
}

// ===================================================================
// 30. Chaining methods that use private fields
// ===================================================================

#[test]
fn method_chaining_with_private_fields() {
    assert_eq!(
        run_js(
            r#"
class Query {
    #table = "";
    #conditions = [];
    #limit = null;
    from(table) { this.#table = table; return this; }
    where(cond) { this.#conditions.push(cond); return this; }
    limit(n) { this.#limit = n; return this; }
    build() {
        let q = "SELECT * FROM " + this.#table;
        if (this.#conditions.length > 0) {
            q += " WHERE " + this.#conditions.join(" AND ");
        }
        if (this.#limit !== null) {
            q += " LIMIT " + this.#limit;
        }
        return q;
    }
}
const result = new Query()
    .from("users")
    .where("age > 18")
    .where("active = true")
    .limit(10)
    .build();
console.log(result);
"#
        ),
        vec!["SELECT * FROM users WHERE age > 18 AND active = true LIMIT 10"]
    );
}

#[test]
fn test_private_field_access_on_other_object_throws_typeerror() {
    assert_eq!(
        run_js(
            r#"
class Secret {
    #code = 1234;
    readOther(obj) {
        return obj.#code;
    }
}
const s = new Secret();
try {
    s.readOther({});
} catch (e) {
    console.log(e.name);
}
"#
        ),
        vec!["TypeError"]
    );
}

#[test]
fn test_private_field_access_on_null_throws_typeerror() {
    assert_eq!(
        run_js(
            r#"
class Secret {
    #code = 1234;
    readNull(obj) {
        return obj.#code;
    }
}
const s = new Secret();
try {
    s.readNull(null);
} catch (e) {
    console.log(e.name);
}
"#
        ),
        vec!["TypeError"]
    );
}
