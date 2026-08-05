use crate::helpers::run_prints;

#[test]
fn test_default_arguments_simple_scalar() {
    let out = run_prints(
        r#"
        fun greet(name: String = "world"): String = "hi " + name
        fun main() {
            println(greet())
            println(greet("kotlin"))
        }
    "#,
    );
    assert_eq!(out, &["hi world", "hi kotlin"]);
}

#[test]
fn test_default_arguments_in_middle_uses_position_and_name() {
    let out = run_prints(
        r#"
        fun score(base: Int, bonus: Int = 1, penalty: Int = 1): Int {
            return base + bonus - penalty
        }
        fun main() {
            println(score(10))
            println(score(10, 3))
            println(score(10, penalty = 4))
        }
    "#,
    );
    // score(10, penalty = 4) is 10 + 1 - 4 = 7 (real Kotlin agrees).
    assert_eq!(out, &["10", "12", "7"]);
}

#[test]
fn test_default_arguments_boolean_default() {
    let out = run_prints(
        r#"
        fun enabled(flag: Boolean = true): String = if (flag) "yes" else "no"
        fun main() {
            println(enabled())
            println(enabled(false))
        }
    "#,
    );
    assert_eq!(out, &["yes", "no"]);
}

#[test]
fn test_default_arguments_lambda_default() {
    let out = run_prints(
        r#"
        fun apply(v: Int, op: (Int) -> Int = { it + 1 }): Int {
            return op(v)
        }
        fun main() {
            println(apply(4))
            println(apply(4, { it * 2 }))
        }
    "#,
    );
    assert_eq!(out, &["5", "8"]);
}

#[test]
fn test_default_arguments_constructor_defaults() {
    let out = run_prints(
        r#"
        class Box(val value: Int = 1, val label: String = "x")
        fun main() {
            val a = Box()
            val b = Box(2)
            val c = Box(label = "z")
            println(a.value)
            println(a.label)
            println(b.value)
            println(c.label)
        }
    "#,
    );
    assert_eq!(out, &["1", "x", "2", "z"]);
}

#[test]
fn test_default_arguments_method_defaults() {
    let out = run_prints(
        r#"
        class Counter {
            fun inc(value: Int = 1): Int = value
        }
        fun main() {
            val c = Counter()
            println(c.inc())
            println(c.inc(3))
        }
    "#,
    );
    assert_eq!(out, &["1", "3"]);
}

#[test]
fn test_default_arguments_chained_defaults() {
    let out = run_prints(
        r#"
        fun base(a: Int = 1): Int = a
        fun step(value: Int = base(), bonus: Int = 2): Int = value + bonus
        fun main() {
            println(step())
            println(step(4))
            println(step(bonus = 5))
        }
    "#,
    );
    assert_eq!(out, &["3", "6", "6"]);
}

#[test]
fn test_default_arguments_list_default_empty() {
    let out = run_prints(
        r#"
        fun join(items: List<String> = listOf()): String = items.joinToString("-")
        fun main() {
            println("<" + join() + ">")
            println(join(listOf("a", "b")))
        }
    "#,
    );
    assert_eq!(out, &["<>", "a-b"]);
}

#[test]
fn test_default_arguments_array_default() {
    let out = run_prints(
        r#"
        fun values(head: Int, rest: IntArray = intArrayOf(1, 2)): Int {
            var sum = head
            for (v in rest) { sum += v }
            return sum
        }
        fun main() {
            println(values(1))
            println(values(1, intArrayOf(5)))
        }
    "#,
    );
    assert_eq!(out, &["4", "6"]);
}

#[test]
fn test_default_arguments_on_extension_function() {
    let out = run_prints(
        r#"
        fun String.wrap(prefix: String = "<", suffix: String = ">"): String {
            return prefix + this + suffix
        }
        fun main() {
            println("a".wrap())
            println("b".wrap(prefix = "["))
        }
    "#,
    );
    // wrap(prefix = "[") keeps the suffix default ">" (real Kotlin agrees).
    assert_eq!(out, &["<a>", "[b>"]);
}

#[test]
fn test_default_arguments_heterogeneous_types() {
    let out = run_prints(
        r#"
        fun combine(a: Int = 1, b: String = "x", c: Boolean = false): String {
            return a.toString() + b + (if (c) "Y" else "N")
        }
        fun main() {
            println(combine())
            println(combine(2, c = true))
            println(combine(b = "z", a = 3))
        }
    "#,
    );
    assert_eq!(out, &["1xN", "2xY", "3zN"]);
}

#[test]
fn test_default_arguments_default_function_parameter_uses_same_default() {
    let out = run_prints(
        r#"
        fun fallback(v: String = "z"): String = v
        fun wrapper(label: String, value: String = fallback()): String = label + value
        fun main() {
            println(wrapper("a"))
            println(wrapper("a", "b"))
        }
    "#,
    );
    assert_eq!(out, &["az", "ab"]);
}

#[test]
fn test_default_arguments_overload_with_defaults_disambiguates_calls() {
    let out = run_prints(
        r#"
        fun pick(a: Int, b: Int = 2): Int = a + b
        fun pick(a: String): String = a
        fun main() {
            println(pick(3))
            println(pick("x"))
            println(pick(3, 4))
        }
    "#,
    );
    assert_eq!(out, &["5", "x", "7"]);
}

#[test]
fn test_default_arguments_generic_default() {
    let out = run_prints(
        r##"
        fun <T> wrap(value: T, marker: String = "#"): String {
            return marker + value.toString()
        }
        fun main() {
            println(wrap(3))
            println(wrap("a", marker = "@"))
        }
    "##,
    );
    assert_eq!(out, &["#3", "@a"]);
}

#[test]
fn test_default_arguments_when_default_reuses_defaulted_param_later() {
    let out = run_prints(
        r#"
        fun score(base: Int = 2, factor: Int = base * 2): Int = factor
        fun main() {
            println(score())
            println(score(5))
            println(score(1, 9))
        }
    "#,
    );
    assert_eq!(out, &["4", "10", "9"]);
}

#[test]
fn test_default_arguments_in_data_class_methods() {
    let out = run_prints(
        r#"
        data class Box(val value: Int, val label: String = "x")
        fun main() {
            val a = Box(1)
            val b = a.copy(label = "y")
            println(a.label)
            println(b.value)
            println(b.label)
        }
    "#,
    );
    assert_eq!(out, &["x", "1", "y"]);
}

#[test]
fn test_default_arguments_defaulted_boolean_chain() {
    let out = run_prints(
        r#"
        fun flags(a: Boolean = true, b: Boolean = false): String = if (a && !b) "on" else "off"
        fun main() {
            println(flags())
            println(flags(a = false))
            println(flags(b = true))
        }
    "#,
    );
    assert_eq!(out, &["on", "off", "off"]);
}

#[test]
fn test_default_arguments_collection_defaults_dont_share_mutable_reference() {
    let out = run_prints(
        r#"
        fun make(items: MutableList<Int> = mutableListOf(1, 2)): String {
            items.add(3)
            return items.joinToString(":")
        }
        fun main() {
            val a = mutableListOf(9)
            println(make(a))
            println(make())
        }
    "#,
    );
    assert_eq!(out, &["9:3", "1:2:3"]);
}

#[test]
fn test_default_arguments_recursive_default_parameters() {
    let out = run_prints(
        r#"
        fun depth(level: Int, suffix: String = ":") : String {
            return if (level <= 0) "0" else depth(level - 1, suffix) + suffix
        }
        fun main() {
            println(depth(0))
            println(depth(2))
        }
    "#,
    );
    assert_eq!(out, &["0", "0::"]);
}

#[test]
fn test_default_arguments_with_named_defaulted_call() {
    let out = run_prints(
        r#"
        fun emit(prefix: String, text: String = "ok", suffix: String = ""): String = prefix + text + suffix
        fun main() {
            println(emit("<", suffix = ">"))
        }
    "#,
    );
    assert_eq!(out, &["<ok>"]);
}

#[test]
fn test_default_arguments_for_local_function() {
    let out = run_prints(
        r#"
        fun main() {
            fun make(base: Int = 1, extra: Int = 2): Int = base + extra
            println(make())
            println(make(5))
            println(make(extra = 10, base = 1))
        }
    "#,
    );
    assert_eq!(out, &["3", "7", "11"]);
}

#[test]
fn test_default_arguments_over_defaulted_class_methods() {
    let out = run_prints(
        r#"
        class Acc {
            fun add(a: Int, b: Int = 1): Int = a + b
            fun nested(label: String = "L"): String = label
        }
        fun main() {
            val a = Acc()
            println(a.add(3))
            println(a.add(3, 4))
            println(a.nested())
        }
    "#,
    );
    assert_eq!(out, &["4", "7", "L"]);
}

#[test]
fn test_default_arguments_nested_defaulting_in_optional_call_sites() {
    let out = run_prints(
        r#"
        fun format(a: String, b: String = "B", c: String = "C"): String = a + b + c
        fun main() {
            println(format("A"))
            println(format("A", c = "X"))
            println(format("A", "Y", "Z"))
        }
    "#,
    );
    // format("A", c = "X") keeps b's default: "ABX" (real Kotlin agrees).
    assert_eq!(out, &["ABC", "ABX", "AYZ"]);
}

#[test]
fn test_default_arguments_default_value_is_literal_object() {
    let out = run_prints(
        r#"
        fun paint(color: String = "red", opacity: Double = 1.0): String = color + ":" + opacity
        fun main() {
            println(paint())
            println(paint(opacity = 0.5))
        }
    "#,
    );
    assert_eq!(out, &["red:1.0", "red:0.5"]);
}

#[test]
fn test_default_arguments_with_defaulted_lambda_returning_string() {
    let out = run_prints(
        r#"
        fun render(prefix: String, printer: () -> String = { "x" }): String {
            return prefix + printer()
        }
        fun main() {
            println(render("a"))
            println(render("a", { "b" }))
        }
    "#,
    );
    assert_eq!(out, &["ax", "ab"]);
}

#[test]
fn test_default_arguments_int_list_defaults() {
    let out = run_prints(
        r#"
        fun sumAll(values: List<Int> = listOf(1, 2, 3)): Int = values.sum()
        fun main() {
            println(sumAll())
            println(sumAll(listOf(10)))
        }
    "#,
    );
    assert_eq!(out, &["6", "10"]);
}

#[test]
fn test_default_arguments_class_static_like() {
    let out = run_prints(
        r#"
        class Counter {
            companion object {
                fun make(base: Int = 9): Int = base
            }
        }
        fun main() {
            println(Counter.make())
            println(Counter.make(2))
        }
    "#,
    );
    assert_eq!(out, &["9", "2"]);
}

#[test]
fn test_default_arguments_nested_method_chain_with_defaults() {
    let out = run_prints(
        r#"
        class Builder {
            fun stage(value: Int = 1): String = (value * 2).toString()
        }
        fun main() {
            println(Builder().stage())
            println(Builder().stage(4))
        }
    "#,
    );
    assert_eq!(out, &["2", "8"]);
}

#[test]
fn test_default_arguments_defaulted_nullable_value() {
    let out = run_prints(
        r#"
        fun pick(v: String?, fallback: String = "d"): String = v ?: fallback
        fun main() {
            println(pick(null))
            println(pick("x"))
            println(pick("", fallback = "z"))
        }
    "#,
    );
    // pick("") passes a NON-null empty string — the elvis keeps it
    // (real Kotlin agrees).
    assert_eq!(out, &["d", "x", ""]);
}

#[test]
fn test_default_arguments_default_on_local_val_return() {
    let out = run_prints(
        r#"
        fun make(label: String = "x", amount: Int = 3): String {
            return label + amount
        }
        fun main() {
            val base = make()
            val changed = make(label = "y", amount = 1)
            println(base)
            println(changed)
        }
    "#,
    );
    assert_eq!(out, &["x3", "y1"]);
}

#[test]
fn test_default_arguments_in_nested_class_context() {
    let out = run_prints(
        r#"
        class Host {
            fun outer(prefix: String = "p", suffix: String = "s") : String = prefix + suffix
            class Child {
                fun inner(tag: String = "t") : String = tag
            }
        }
        fun main() {
            val h = Host()
            val c = Host.Child()
            println(h.outer())
            println(c.inner())
            println(c.inner("x"))
        }
    "#,
    );
    assert_eq!(out, &["ps", "t", "x"]);
}

#[test]
fn test_default_arguments_trailing_default_not_passed() {
    let out = run_prints(
        r#"
        fun combine(prefix: String, postfix: String = "X", center: String = "Y"): String {
            return prefix + center + postfix
        }
        fun main() {
            println(combine("a"))
            println(combine("a", "B"))
            println(combine("a", center = "C", postfix = "D"))
        }
    "#,
    );
    assert_eq!(out, &["aYX", "aYB", "aCD"]);
}

#[test]
fn test_default_arguments_method_reference_keeps_defaults() {
    let out = run_prints(
        r##"
        fun decorate(text: String, marker: String = "*"): String = marker + text + marker
        fun main() {
            val f = ::decorate
            println(f("x"))
            println(f("y", "#"))
        }
    "##,
    );
    assert_eq!(out, &["*x*", "#y#"]);
}

#[test]
fn test_default_arguments_nested_defaults_with_defaults_on_named_type() {
    let out = run_prints(
        r#"
        fun line(a: String = "a", b: String = a): String = a + b
        fun main() {
            println(line())
            println(line("x"))
        }
    "#,
    );
    assert_eq!(out, &["aa", "xx"]);
}

#[test]
fn test_default_arguments_finality_with_many_params() {
    let out = run_prints(
        r#"
        fun eval(a: Int, b: Int = 1, c: Int = 2, d: Int = 3): Int {
            return a + b + c + d
        }
        fun main() {
            println(eval(1))
            println(eval(1, d = 10))
            println(eval(1, 2, 3, 4))
        }
    "#,
    );
    // eval(1, d = 10) is 1 + 1 + 2 + 10 = 14 (real Kotlin agrees).
    assert_eq!(out, &["7", "14", "10"]);
}
