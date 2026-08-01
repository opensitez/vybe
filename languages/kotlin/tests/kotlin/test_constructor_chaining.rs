use crate::helpers::run_prints;

#[test]
fn test_constructor_primary_simple() {
    let out = run_prints(
        r#"
        class Box(val v: Int)
        fun main() {
            println(Box(1).v)
        }
    "#,
    );
    assert_eq!(out, &["1"]);
}

#[test]
fn test_constructor_secondary_default() {
    let out = run_prints(
        r#"
        class Box {
            val v: Int
            constructor(v: Int) { this.v = v }
            constructor() : this(3)
        }
        fun main() {
            println(Box().v)
            println(Box(5).v)
        }
    "#,
    );
    assert_eq!(out, &["3", "5"]);
}

#[test]
fn test_constructor_inheritance_chain() {
    let out = run_prints(
        r#"
        open class A(val a: Int)
        class B(x: Int) : A(x + 1)
        fun main() {
            println(B(2).a)
        }
    "#,
    );
    assert_eq!(out, &["3"]);
}

#[test]
fn test_constructor_with_init() {
    let out = run_prints(
        r#"
        class Track(val x: Int) {
            val y: Int
            init { y = x * 2 }
        }
        fun main() {
            println(Track(4).y)
        }
    "#,
    );
    assert_eq!(out, &["8"]);
}

#[test]
fn test_constructor_init_read_order() {
    let out = run_prints(
        r#"
        class Bag(val base: Int) {
            val a: Int
            init { a = base + 1 }
            val b = a + 2
        }
        fun main() {
            val b = Bag(3)
            println(b.a)
            println(b.b)
        }
    "#,
    );
    assert_eq!(out, &["4", "6"]);
}

#[test]
fn test_constructor_private_constructor() {
    let out = run_prints(
        r#"
        class Secret private constructor(val v: Int) {
            companion object {
                fun create(v: Int) = Secret(v)
            }
        }
        fun main() {
            println(Secret.create(9).v)
        }
    "#,
    );
    assert_eq!(out, &["9"]);
}

#[test]
fn test_constructor_secondary_chain_to_primary() {
    let out = run_prints(
        r#"
        class Chain {
            val v: Int
            constructor() { this(7) }
            constructor(v: Int) { this.v = v }
        }
        fun main() {
            println(Chain().v)
        }
    "#,
    );
    assert_eq!(out, &["7"]);
}

#[test]
fn test_constructor_nested_defaults() {
    let out = run_prints(
        r#"
        class C(val a: Int, val b: Int = 2) {
            constructor() : this(0)
            constructor(a: Int, b: Int, c: Int) : this(a + b + c)
        }
        fun main() {
            println(C().b)
            println(C(1, 2).a)
            println(C(1, 2, 3).a)
        }
    "#,
    );
    assert_eq!(out, &["2", "1", "6"]);
}

#[test]
fn test_constructor_generic_class() {
    let out = run_prints(
        r#"
        class Holder<T>(val v: T) {
            val text = v.toString()
        }
        fun main() {
            val h = Holder("x")
            println(h.text)
        }
    "#,
    );
    assert_eq!(out, &["x"]);
}

#[test]
fn test_constructor_data_class_copy_ctor() {
    let out = run_prints(
        r#"
        data class P(val a: Int, val b: String)
        fun main() {
            val p = P(1, "x")
            val c = p.copy(a = 2)
            println(c.a)
            println(c.b)
        }
    "#,
    );
    assert_eq!(out, &["2", "x"]);
}

#[test]
fn test_constructor_with_interface_impl() {
    let out = run_prints(
        r#"
        interface I { fun tag(): String }
        class C(val v: Int) : I {
            override fun tag() = "c:$v"
        }
        fun main() {
            println(C(7).tag())
        }
    "#,
    );
    assert_eq!(out, &["c:7"]);
}

#[test]
fn test_constructor_with_companion_factory() {
    let out = run_prints(
        r#"
        class PairNum private constructor(val a: Int, val b: Int) {
            companion object {
                fun of(a: Int) = PairNum(a, a + 1)
            }
        }
        fun main() {
            val p = PairNum.of(2)
            println(p.a + p.b)
        }
    "#,
    );
    assert_eq!(out, &["5"]);
}

#[test]
fn test_constructor_with_named_arg_defaults() {
    let out = run_prints(
        r#"
        class Namer(val first: String = "x", val last: String = "y")
        fun main() {
            val a = Namer(last = "z")
            println(a.first)
            println(a.last)
        }
    "#,
    );
    assert_eq!(out, &["x", "z"]);
}

#[test]
fn test_constructor_multiple_init() {
    let out = run_prints(
        r#"
        class Multi {
            val a: Int
            init {
                println("init")
                a = 1
            }
        }
        fun main() {
            val m = Multi()
            println(m.a)
        }
    "#,
    );
    assert_eq!(out, &["init", "1"]);
}

#[test]
fn test_constructor_inherited_override_property() {
    let out = run_prints(
        r#"
        open class Base { open val v = 1 }
        class Sub : Base() { override val v = 2 }
        fun main() {
            val b: Base = Sub()
            println(b.v)
        }
    "#,
    );
    assert_eq!(out, &["2"]);
}

#[test]
fn test_constructor_with_default_lambda() {
    let out = run_prints(
        r#"
        class Maker(val f: () -> Int = { 3 }) {
            fun value() = f()
        }
        fun main() {
            println(Maker().value())
            println(Maker { 5 } .value())
        }
    "#,
    );
    assert_eq!(out, &["3", "5"]);
}

#[test]
fn test_constructor_captures_param_in_init() {
    let out = run_prints(
        r#"
        class Capture(val x: Int) {
            val y: Int
            init { y = x * 3 }
        }
        fun main() {
            println(Capture(4).y)
        }
    "#,
    );
    assert_eq!(out, &["12"]);
}

#[test]
fn test_constructor_secondary_calls_base_super() {
    let out = run_prints(
        r#"
        open class Parent(val a: Int)
        class Child : Parent {
            constructor() : super(1)
        }
        fun main() {
            println(Child().a)
        }
    "#,
    );
    assert_eq!(out, &["1"]);
}

#[test]
fn test_constructor_chained_via_this() {
    let out = run_prints(
        r#"
        class ThisChain {
            val v: Int
            constructor() : this(8)
            constructor(v: Int) { this.v = v }
        }
        fun main() {
            println(ThisChain().v)
        }
    "#,
    );
    assert_eq!(out, &["8"]);
}

#[test]
fn test_constructor_array_wrapping() {
    let out = run_prints(
        r#"
        class Arr(val items: IntArray)
        fun main() {
            val a = Arr(intArrayOf(1, 2, 3))
            println(a.items[1])
        }
    "#,
    );
    assert_eq!(out, &["2"]);
}

#[test]
fn test_constructor_secondary_boolean() {
    let out = run_prints(
        r#"
        class Flag {
            val enabled: Boolean
            constructor(enabled: Boolean) { this.enabled = enabled }
            constructor() : this(false)
        }
        fun main() {
            println(Flag().enabled)
            println(Flag(true).enabled)
        }
    "#,
    );
    assert_eq!(out, &["false", "true"]);
}

#[test]
fn test_constructor_reified_like_no() {
    let out = run_prints(
        r#"
        class Holder {
            val value: String
            constructor(v: Int) { value = v.toString() }
            constructor(v: String) { value = v }
        }
        fun main() {
            println(Holder(4).value)
            println(Holder("x").value)
        }
    "#,
    );
    assert_eq!(out, &["4", "x"]);
}

#[test]
fn test_constructor_property_order() {
    let out = run_prints(
        r#"
        class Order {
            val a: Int
            val b: Int
            init { b = 2; a = 1 }
        }
        fun main() {
            val o = Order()
            println(o.a)
            println(o.b)
        }
    "#,
    );
    assert_eq!(out, &["1", "2"]);
}

#[test]
fn test_constructor_with_default_object_expr() {
    let out = run_prints(
        r#"
        class Holder(val value: Int)
        fun main() {
            val h: Holder = Holder(3)
            println(h.value)
        }
    "#,
    );
    assert_eq!(out, &["3"]);
}

#[test]
fn test_constructor_init_side_effect_print() {
    let out = run_prints(
        r#"
        class Side {
            init { println("ok") }
        }
        fun main() {
            Side()
        }
    "#,
    );
    assert_eq!(out, &["ok"]);
}

#[test]
fn test_constructor_property_with_expr_init() {
    let out = run_prints(
        r#"
        class Calc(val a: Int) {
            val b = a + 1
            val c = b + a
        }
        fun main() {
            val c = Calc(4)
            println(c.b)
            println(c.c)
        }
    "#,
    );
    assert_eq!(out, &["5", "9"]);
}

#[test]
fn test_constructor_nested_class_call() {
    let out = run_prints(
        r#"
        class Outer {
            class Inner(val v: Int)
        }
        fun main() {
            val i = Outer.Inner(9)
            println(i.v)
        }
    "#,
    );
    assert_eq!(out, &["9"]);
}

#[test]
fn test_constructor_overload_count() {
    let out = run_prints(
        r#"
        class Over {
            val label: String
            constructor() { label = "x" }
            constructor(v: Int) { label = v.toString() }
        }
        fun main() {
            println(Over().label)
            println(Over(2).label)
        }
    "#,
    );
    assert_eq!(out, &["x", "2"]);
}

#[test]
fn test_constructor_mismatched_types_compile_path() {
    let out = run_prints(
        r#"
        class Num(val text: String) {
            constructor(v: Int) : this(v.toString())
        }
        fun main() {
            println(Num(10).text)
        }
    "#,
    );
    assert_eq!(out, &["10"]);
}

#[test]
fn test_constructor_in_default_body() {
    let out = run_prints(
        r#"
        class DefaultBody {
            val value: Int
            constructor(v: Int = 1) { value = v }
        }
        fun main() {
            println(DefaultBody().value)
        }
    "#,
    );
    assert_eq!(out, &["1"]);
}

#[test]
fn test_constructor_two_secondary_calls() {
    let out = run_prints(
        r#"
        class Layer {
            val a: Int
            val b: Int
            constructor(a: Int) { this.a = a; this.b = a }
            constructor(a: Int, b: Int) : this(a + b)
        }
        fun main() {
            val l = Layer(2, 3)
            println(l.a)
            println(l.b)
        }
    "#,
    );
    assert_eq!(out, &["5", "5"]);
}

#[test]
fn test_constructor_with_property_param_override() {
    let out = run_prints(
        r#"
        open class Base(val v: Int)
        class Derived(v: Int) : Base(v + 1)
        fun main() {
            println(Derived(4).v)
        }
    "#,
    );
    assert_eq!(out, &["5"]);
}

#[test]
fn test_constructor_with_boolean_expression() {
    let out = run_prints(
        r#"
        class Tri(val v: Int) {
            val ok: Boolean
            init { ok = v % 2 == 0 }
        }
        fun main() {
            println(Tri(4).ok)
            println(Tri(5).ok)
        }
    "#,
    );
    assert_eq!(out, &["true", "false"]);
}

#[test]
fn test_constructor_final_read() {
    let out = run_prints(
        r#"
        class FinalRead {
            val first = 1
        }
        fun main() {
            println(FinalRead().first)
        }
    "#,
    );
    assert_eq!(out, &["1"]);
}

#[test]
fn test_constructor_list_param() {
    let out = run_prints(
        r#"
        class WithList(val items: List<Int>)
        fun main() {
            val w = WithList(listOf(1, 2))
            println(w.items.joinToString("|"))
        }
    "#,
    );
    assert_eq!(out, &["1|2"]);
}
