use super::helpers::run_js;

fn run_one(code: &str) -> String {
    run_js(code).into_iter().next().unwrap_or_default()
}

// ── Map tests ──────────────────────────────────────────────

#[test]
fn map_set_get() {
    assert_eq!(
        run_one(
            r#"
        let m = new Map();
        m.set("name", "Alice");
        console.log(m.get("name"));
    "#
        ),
        "Alice"
    );
}

#[test]
fn map_has() {
    assert_eq!(
        run_one(
            r#"
        let m = new Map();
        m.set("x", 1);
        console.log(m.has("x"), m.has("y"));
    "#
        ),
        "true false"
    );
}

#[test]
fn map_size() {
    assert_eq!(
        run_one(
            r#"
        let m = new Map();
        m.set("a", 1);
        m.set("b", 2);
        console.log(m.size);
    "#
        ),
        "2"
    );
}

#[test]
fn map_overwrite_no_size_change() {
    assert_eq!(
        run_one(
            r#"
        let m = new Map();
        m.set("a", 1);
        m.set("a", 2);
        console.log(m.size, m.get("a"));
    "#
        ),
        "1 2"
    );
}

#[test]
fn map_delete() {
    assert_eq!(
        run_one(
            r#"
        let m = new Map();
        m.set("a", 1);
        m.set("b", 2);
        m.delete("a");
        console.log(m.size, m.has("a"), m.get("b"));
    "#
        ),
        "1 false 2"
    );
}

#[test]
fn map_clear() {
    assert_eq!(
        run_one(
            r#"
        let m = new Map();
        m.set("a", 1);
        m.set("b", 2);
        m.clear();
        console.log(m.size);
    "#
        ),
        "0"
    );
}

#[test]
fn map_keys() {
    let out = run_js(
        r#"
        let m = new Map();
        m.set("x", 10);
        m.set("y", 20);
        let keys = m.keys();
        console.log(keys.length);
    "#,
    );
    assert_eq!(out[0], "undefined");
}

#[test]
fn map_chaining() {
    assert_eq!(
        run_one(
            r#"
        let m = new Map();
        m.set("a", 1).set("b", 2).set("c", 3);
        console.log(m.size);
    "#
        ),
        "3"
    );
}

#[test]
fn map_operations_sequence() {
    assert_eq!(
        run_one(
            r#"
        let m = new Map();
        m.set("a", 1);
        m.set("b", 2);
        m.set("c", 3);
        m.delete("b");
        m.set("d", 4);
        console.log(m.size, m.has("a"), m.has("b"), m.get("d"));
    "#
        ),
        "3 true false 4"
    );
}

// ── Set tests ──────────────────────────────────────────────

#[test]
fn set_add_has() {
    assert_eq!(
        run_one(
            r#"
        let s = new Set();
        s.add("hello");
        console.log(s.has("hello"), s.has("world"));
    "#
        ),
        "true false"
    );
}

#[test]
fn set_size() {
    assert_eq!(
        run_one(
            r#"
        let s = new Set();
        s.add(1);
        s.add(2);
        s.add(3);
        console.log(s.size);
    "#
        ),
        "3"
    );
}

#[test]
fn set_no_duplicates() {
    assert_eq!(
        run_one(
            r#"
        let s = new Set();
        s.add(1);
        s.add(2);
        s.add(1);
        s.add(2);
        s.add(3);
        console.log(s.size);
    "#
        ),
        "3"
    );
}

#[test]
fn set_delete() {
    assert_eq!(
        run_one(
            r#"
        let s = new Set();
        s.add("a");
        s.add("b");
        s.delete("a");
        console.log(s.size, s.has("a"));
    "#
        ),
        "1 false"
    );
}

#[test]
fn set_clear() {
    assert_eq!(
        run_one(
            r#"
        let s = new Set();
        s.add(1);
        s.add(2);
        s.clear();
        console.log(s.size);
    "#
        ),
        "0"
    );
}

// ── Method chaining ────────────────────────────────────────

#[test]
fn builder_chain() {
    assert_eq!(
        run_one(
            r#"
        class Builder {
            constructor() { this.parts = []; }
            add(s) { this.parts.push(s); return this; }
            build() { return this.parts.join("-"); }
        }
        let b = new Builder();
        console.log(b.add("a").add("b").add("c").build());
    "#
        ),
        "a-b-c"
    );
}

#[test]
fn chain_two() {
    assert_eq!(
        run_one(
            r#"
        class B {
            constructor() { this.parts = []; }
            add(s) { this.parts.push(s); return this; }
            build() { return this.parts.join("-"); }
        }
        let b = new B();
        console.log(b.add("x").build());
    "#
        ),
        "x"
    );
}

#[test]
fn chain_separate_calls() {
    assert_eq!(
        run_one(
            r#"
        class B {
            constructor() { this.parts = []; }
            add(s) { this.parts.push(s); return this; }
            build() { return this.parts.join("-"); }
        }
        let b = new B();
        b.add("a");
        b.add("b");
        b.add("c");
        console.log(b.build());
    "#
        ),
        "a-b-c"
    );
}

// ── Closure on object methods ──────────────────────────────

#[test]
fn closure_read_on_object() {
    assert_eq!(
        run_one(
            r#"
        function make() { let x = 99; return { getValue: () => { return x; } }; }
        let o = make();
        console.log(o.getValue());
    "#
        ),
        "99"
    );
}

#[test]
fn closure_mutate_on_object() {
    assert_eq!(
        run_one(
            r#"
        function make() {
            let n = 0;
            return {
                inc: () => { n = n + 1; return n; },
                getN: () => n
            };
        }
        let c = make();
        c.inc();
        c.inc();
        c.inc();
        console.log(c.getN());
    "#
        ),
        "3"
    );
}

// ── Class with Map field ───────────────────────────────────

#[test]
fn class_map_field() {
    assert_eq!(
        run_one(
            r#"
        class Registry {
            constructor() { this.data = new Map(); }
            register(key, val) { this.data.set(key, val); }
            lookup(key) { return this.data.get(key); }
        }
        let r = new Registry();
        r.register("host", "localhost");
        console.log(r.lookup("host"));
    "#
        ),
        "localhost"
    );
}

#[test]
fn map_key_negative_zero_same_as_positive_zero() {
    assert_eq!(
        run_one(
            r#"
        let m = new Map();
        m.set(-0, "zero");
        console.log(m.get(+0), m.has(+0));
    "#
        ),
        "zero true"
    );
}

