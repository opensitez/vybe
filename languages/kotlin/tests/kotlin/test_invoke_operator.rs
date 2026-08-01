use crate::helpers::run_prints;

#[test]
fn test_invoke_no_args() {
    let out = run_prints(
        r#"
        class Counter {
            operator fun invoke(): Int = 1
        }
        fun main() {
            println(Counter()())
        }
    "#,
    );
    assert_eq!(out, &["1"]);
}

#[test]
fn test_invoke_one_arg() {
    let out = run_prints(
        r#"
        class Math {
            operator fun invoke(v: Int): Int = v + 1
        }
        fun main() {
            println(Math()(3))
        }
    "#,
    );
    assert_eq!(out, &["4"]);
}

#[test]
fn test_invoke_two_args() {
    let out = run_prints(
        r#"
        class PairAdder {
            operator fun invoke(a: Int, b: Int): Int = a + b
        }
        fun main() {
            println(PairAdder()(2, 5))
        }
    "#,
    );
    assert_eq!(out, &["7"]);
}

#[test]
fn test_invoke_three_args() {
    let out = run_prints(
        r#"
        class Sum3 {
            operator fun invoke(a: Int, b: Int, c: Int): Int = a + b + c
        }
        fun main() {
            println(Sum3()(1, 2, 3))
        }
    "#,
    );
    assert_eq!(out, &["6"]);
}

#[test]
fn test_invoke_infix_style_variable() {
    let out = run_prints(
        r#"
        class Box {
            operator fun invoke(v: String): String = "[$v]"
        }
        fun main() {
            val f = Box()
            println(f("x"))
        }
    "#,
    );
    assert_eq!(out, &["[x]"]);
}

#[test]
fn test_invoke_via_property() {
    let out = run_prints(
        r#"
        class Builder {
            val call = { n: Int -> n * 2 }
            operator fun invoke(v: Int): Int = call(v)
        }
        fun main() {
            val b = Builder()
            println(b(4))
        }
    "#,
    );
    assert_eq!(out, &["8"]);
}

#[test]
fn test_invoke_with_nullable_receiver() {
    let out = run_prints(
        r#"
        class Greeter {
            operator fun invoke(name: String?): String = name ?: "guest"
        }
        fun main() {
            val g = Greeter()
            println(g(null))
        }
    "#,
    );
    assert_eq!(out, &["guest"]);
}

#[test]
fn test_invoke_in_class_inheritance() {
    let out = run_prints(
        r#"
        open class Base {
            operator fun invoke(v: Int): Int = v
        }
        class Child : Base()
        fun main() {
            println(Child()(9))
        }
    "#,
    );
    assert_eq!(out, &["9"]);
}

#[test]
fn test_invoke_from_local_object() {
    let out = run_prints(
        r#"
        fun main() {
            val f = object {
                operator fun invoke(v: Int): Int = v * v
            }
            println(f(6))
        }
    "#,
    );
    assert_eq!(out, &["36"]);
}

#[test]
fn test_invoke_overload_different_types() {
    let out = run_prints(
        r#"
        class Printer {
            operator fun invoke(v: Int): String = "i$v"
            operator fun invoke(v: String): String = "s$v"
        }
        fun main() {
            val p = Printer()
            println(p(3))
            println(p("x"))
        }
    "#,
    );
    assert_eq!(out, &["i3", "sx"]);
}

#[test]
fn test_invoke_extension_function_type() {
    let out = run_prints(
        r#"
        class Host {
            operator fun String.invoke(v: String): String = this + v
        }
        fun main() {
            val host = Host()
            with(host) {
                println("a".invoke("b"))
            }
        }
    "#,
    );
    assert_eq!(out, &["ab"]);
}

#[test]
fn test_invoke_with_defaulted_parameter() {
    let out = run_prints(
        r#"
        class Formatter {
            operator fun invoke(prefix: String = "[") : String = prefix + "end]"
        }
        fun main() {
            val f = Formatter()
            println(f())
            println(f("(") )
        }
    "#,
    );
    assert_eq!(out, &["[end]", "(end]"]);
}

#[test]
fn test_invoke_on_function_reference() {
    let out = run_prints(
        r#"
        class Adder {
            operator fun invoke(x: Int): Int = x + 1
        }
        fun main() {
            val fn: (Int) -> Int = Adder()
            println(fn(2))
        }
    "#,
    );
    assert_eq!(out, &["3"]);
}

#[test]
fn test_invoke_chainable() {
    let out = run_prints(
        r#"
        class A {
            operator fun invoke(x: Int): B = B(x + 1)
        }
        class B(val v: Int) {
            operator fun invoke(y: Int): Int = v + y
        }
        fun main() {
            println(A()(3)(4))
        }
    "#,
    );
    assert_eq!(out, &["8"]);
}

#[test]
fn test_invoke_operator_variance_through_type_alias() {
    let out = run_prints(
        r#"
        typealias IntCall = (Int) -> Int
        class Wrapper {
            operator fun invoke(v: Int): IntCall = { it + v }
        }
        fun main() {
            val f = Wrapper()(3)
            println(f(7))
        }
    "#,
    );
    assert_eq!(out, &["10"]);
}

#[test]
fn test_invoke_inside_lambda_argument() {
    let out = run_prints(
        r#"
        class Twice {
            operator fun invoke(v: Int): Int = v * 2
        }
        fun run(v: Int, fn: (Int) -> Int): Int = fn(v)
        fun main() {
            println(run(5, Twice()))
        }
    "#,
    );
    assert_eq!(out, &["10"]);
}

#[test]
fn test_invoke_boolean_toggle() {
    let out = run_prints(
        r#"
        class Toggle {
            var state = false
            operator fun invoke(): Boolean {
                state = !state
                return state
            }
        }
        fun main() {
            val t = Toggle()
            println(t())
            println(t())
        }
    "#,
    );
    assert_eq!(out, &["true", "false"]);
}

#[test]
fn test_invoke_string_transformer() {
    let out = run_prints(
        r#"
        class Repeater {
            operator fun invoke(v: String, n: Int): String = v.repeat(n)
        }
        fun main() {
            println(Repeater()("a", 3))
        }
    "#,
    );
    assert_eq!(out, &["aaa"]);
}

#[test]
fn test_invoke_using_this_reference() {
    let out = run_prints(
        r#"
        class Counter {
            private var total = 0
            operator fun invoke(): Int {
                total += 1
                return total
            }
        }
        fun main() {
            val c = Counter()
            println(c())
            println(c())
            println(c())
        }
    "#,
    );
    assert_eq!(out, &["1", "2", "3"]);
}

#[test]
fn test_invoke_nested_object_expression() {
    let out = run_prints(
        r#"
        fun main() {
            val factory = object {
                operator fun invoke(a: Int, b: Int): Int = a + b
            }
            println(factory(4, 5))
        }
    "#,
    );
    assert_eq!(out, &["9"]);
}

#[test]
fn test_invoke_with_vararg() {
    let out = run_prints(
        r#"
        class Summer {
            operator fun invoke(vararg values: Int): Int = values.sum()
        }
        fun main() {
            println(Summer()(1, 2, 3))
        }
    "#,
    );
    assert_eq!(out, &["6"]);
}

#[test]
fn test_invoke_array_argument() {
    let out = run_prints(
        r#"
        class Adder {
            operator fun invoke(values: IntArray): Int = values.sum()
        }
        fun main() {
            println(Adder()(intArrayOf(3, 4, 5)))
        }
    "#,
    );
    assert_eq!(out, &["12"]);
}

#[test]
fn test_invoke_list_argument() {
    let out = run_prints(
        r#"
        class Joiner {
            operator fun invoke(parts: List<String>): String = parts.joinToString(":")
        }
        fun main() {
            println(Joiner()(listOf("a", "b")))
        }
    "#,
    );
    assert_eq!(out, &["a:b"]);
}

#[test]
fn test_invoke_unit_return() {
    let out = run_prints(
        r#"
        class Sink {
            private var logged = false
            operator fun invoke(v: Int): Unit { logged = v > 0 }
            fun status() = logged
        }
        fun main() {
            val s = Sink()
            s(3)
            println(s.status())
            s(-1)
            println(s.status())
        }
    "#,
    );
    assert_eq!(out, &["true", "false"]);
}

#[test]
fn test_invoke_boolean_predicate() {
    let out = run_prints(
        r#"
        class Matcher {
            operator fun invoke(v: Int): Boolean = v % 2 == 0
        }
        fun main() {
            val m = Matcher()
            println(m(4))
            println(m(5))
        }
    "#,
    );
    assert_eq!(out, &["true", "false"]);
}

#[test]
fn test_invoke_as_map_lookup() {
    let out = run_prints(
        r#"
        class Lookup {
            private val data = mapOf("a" to 1, "b" to 2)
            operator fun invoke(key: String): Int = data[key] ?: 0
        }
        fun main() {
            val l = Lookup()
            println(l("a"))
            println(l("z"))
        }
    "#,
    );
    assert_eq!(out, &["1", "0"]);
}

#[test]
fn test_invoke_mutable_state_counter() {
    let out = run_prints(
        r#"
        class Meter {
            private var total = 0
            operator fun invoke(n: Int) {
                total += n
            }
            fun value(): Int = total
        }
        fun main() {
            val m = Meter()
            m(3)
            m(4)
            println(m.value())
        }
    "#,
    );
    assert_eq!(out, &["7"]);
}

#[test]
fn test_invoke_lambda_wrapping() {
    let out = run_prints(
        r#"
        class Wrapper {
            operator fun invoke(v: (Int) -> Int): Int = v(6)
        }
        fun main() {
            val w = Wrapper()
            println(w { it * 2 })
        }
    "#,
    );
    assert_eq!(out, &["12"]);
}

#[test]
fn test_invoke_no_args_with_side_effect() {
    let out = run_prints(
        r#"
        var count = 0
        class Notifier {
            operator fun invoke() {
                count += 1
            }
        }
        fun main() {
            val n = Notifier()
            n()
            n()
            println(count)
        }
    "#,
    );
    assert_eq!(out, &["2"]);
}

#[test]
fn test_invoke_operator_dispatch_in_subclass() {
    let out = run_prints(
        r#"
        open class A {
            open operator fun invoke(v: String): String = "A: " + v
        }
        class B : A() {
            override operator fun invoke(v: String): String = "B: " + v
        }
        fun main() {
            val a: A = B()
            println(a("x"))
        }
    "#,
    );
    assert_eq!(out, &["B: x"]);
}

#[test]
fn test_invoke_operator_overload_heterogeneous_types() {
    let out = run_prints(
        r#"
        class C {
            operator fun invoke(v: Int): String = "i$v"
            operator fun invoke(v: Double): String = "d$v"
            operator fun invoke(v: Boolean): String = if (v) "on" else "off"
        }
        fun main() {
            val c = C()
            println(c(1))
            println(c(2.5))
            println(c(false))
        }
    "#,
    );
    assert_eq!(out, &["i1", "d2.5", "off"]);
}

#[test]
fn test_invoke_variadic_with_named_style_notation() {
    let out = run_prints(
        r#"
        class Tagger {
            operator fun invoke(prefix: String, value: String = "x"): String = prefix + value
        }
        fun main() {
            val t = Tagger()
            println(t("a"))
            println(t(prefix = "b", value = "y"))
        }
    "#,
    );
    assert_eq!(out, &["ax", "by"]);
}

#[test]
fn test_invoke_in_when_dispatch() {
    let out = run_prints(
        r#"
        class Router {
            operator fun invoke(flag: Boolean): String = if (flag) "ok" else "bad"
        }
        fun main() {
            val r = Router()
            val a = true
            val b = false
            println(r(a))
            println(r(b))
        }
    "#,
    );
    assert_eq!(out, &["ok", "bad"]);
}

#[test]
fn test_invoke_tail_style_reuse() {
    let out = run_prints(
        r#"
        class Tail {
            operator fun invoke(v: Int): Tail {
                return if (v <= 0) this else Tail()
            }
            val id: Int = 1
        }
        fun main() {
            val t = Tail()
            println(t(0).id)
            println(t(1).id)
        }
    "#,
    );
    assert_eq!(out, &["1", "1"]);
}
