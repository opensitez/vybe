use crate::helpers::run_prints;

#[test]
fn test_function_type_basic_inference() {
    let out = run_prints(
        r#"
        fun main() {
            val f: (Int) -> Int = { it + 1 }
            println(f(2))
        }
    "#,
    );
    assert_eq!(out, &["3"]);
}

#[test]
fn test_function_type_as_return_type() {
    let out = run_prints(
        r#"
        fun maker(): (Int) -> Int {
            return { v -> v * v }
        }
        fun main() {
            println(maker()(3))
        }
    "#,
    );
    assert_eq!(out, &["9"]);
}

#[test]
fn test_function_type_parameter_named() {
    let out = run_prints(
        r#"
        fun apply(value: Int, op: (Int) -> Int): Int {
            return op(value)
        }
        fun main() {
            println(apply(4, { it * 3 }))
            println(apply(4) { it + 1 })
        }
    "#,
    );
    assert_eq!(out, &["12", "5"]);
}

#[test]
fn test_function_type_with_two_args() {
    let out = run_prints(
        r#"
        fun combine(a: Int, b: Int, fn: (Int, Int) -> Int): Int {
            return fn(a, b)
        }
        fun main() {
            println(combine(2, 4, { x, y -> x + y }))
            println(combine(2, 4, Int::plus))
        }
    "#,
    );
    assert_eq!(out, &["6", "6"]);
}

#[test]
fn test_function_type_nullable() {
    let out = run_prints(
        r#"
        fun handle(fn: ((Int) -> Int)?): Int {
            return fn?.invoke(4) ?: 0
        }
        fun main() {
            println(handle(null))
            println(handle({ it + 1 }))
        }
    "#,
    );
    assert_eq!(out, &["0", "5"]);
}

#[test]
fn test_function_type_extension_receiver() {
    let out = run_prints(
        r#"
        fun main() {
            val upper: String.() -> String = { uppercase() }
            println("a".upper())
            val append: String.(String) -> String = { this + it }
            println("x".append("y"))
        }
    "#,
    );
    assert_eq!(out, &["A", "xy"]);
}

#[test]
fn test_function_type_nested_function() {
    let out = run_prints(
        r#"
        fun main() {
            fun produce(): (String) -> Int {
                val base = 1
                return { it.length + base }
            }
            val f = produce()
            println(f("ab"))
        }
    "#,
    );
    assert_eq!(out, &["3"]);
}

#[test]
fn test_function_type_alias_basic() {
    let out = run_prints(
        r#"
        typealias IntOp = (Int) -> Int
        fun transform(v: Int, op: IntOp): Int = op(v)
        fun main() {
            val square: IntOp = { it * it }
            println(transform(5, square))
        }
    "#,
    );
    assert_eq!(out, &["25"]);
}

#[test]
fn test_function_reference_to_top_level() {
    let out = run_prints(
        r#"
        fun inc(v: Int): Int = v + 1
        fun main() {
            val f: (Int) -> Int = ::inc
            println(f(2))
        }
    "#,
    );
    assert_eq!(out, &["3"]);
}

#[test]
fn test_function_reference_member() {
    let out = run_prints(
        r##"
        class A {
            fun show(v: Int): String = "#" + v
        }
        fun main() {
            val a = A()
            val f = a::show
            println(f(7))
        }
    "##,
    );
    assert_eq!(out, &["#7"]);
}

#[test]
fn test_function_type_stored_in_collection() {
    let out = run_prints(
        r#"
        fun main() {
            val ops: List<(Int) -> Int> = listOf({ it + 1 }, { it * 2 }, { it * it })
            println(ops[0](3))
            println(ops[1](3))
            println(ops[2](3))
        }
    "#,
    );
    assert_eq!(out, &["4", "6", "9"]);
}

#[test]
fn test_function_type_higher_order_chain() {
    let out = run_prints(
        r#"
        fun map(v: Int, first: (Int) -> Int, second: (Int) -> Int): Int {
            return second(first(v))
        }
        fun main() {
            val out = map(2, { it + 3 }, { it * 4 })
            println(out)
        }
    "#,
    );
    assert_eq!(out, &["20"]);
}

#[test]
fn test_function_type_with_default_argument() {
    let out = run_prints(
        r#"
        fun run(value: Int, op: (Int) -> Int = { it + 1 }): Int = op(value)
        fun main() {
            println(run(2))
            println(run(2, { it * 3 }))
        }
    "#,
    );
    assert_eq!(out, &["3", "6"]);
}

#[test]
fn test_function_type_of_unit() {
    let out = run_prints(
        r#"
        fun apply(side: (Int) -> Unit): String {
            side(3)
            return "ok"
        }
        fun main() {
            println(apply({ println("x" + it); }))
        }
    "#,
    );
    assert_eq!(out, &["x3", "ok"]);
}

#[test]
fn test_function_type_as_result_of_else_branch() {
    let out = run_prints(
        r#"
        fun pick(upper: Boolean): (Int) -> Int {
            return if (upper) { { it * 2 } } else { { it + 5 } }
        }
        fun main() {
            println(pick(true)(3))
            println(pick(false)(3))
        }
    "#,
    );
    assert_eq!(out, &["6", "8"]);
}

#[test]
fn test_function_type_with_trailing_lambda_call_style() {
    let out = run_prints(
        r#"
        fun use(v: Int, block: (Int) -> Int): Int = block(v)
        fun main() {
            println(use(4) { it + 10 })
        }
    "#,
    );
    assert_eq!(out, &["14"]);
}

#[test]
fn test_function_type_with_boolean_predicate_list() {
    let out = run_prints(
        r#"
        fun filter(values: List<Int>, keep: (Int) -> Boolean): List<Int> {
            return values.filter(keep)
        }
        fun main() {
            val out = filter(listOf(1, 2, 3, 4), { it % 2 == 0 })
            println(out.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["2,4"]);
}

#[test]
fn test_function_type_in_data_flow() {
    let out = run_prints(
        r#"
        fun main() {
            val a: Int = 1
            val f: (Int) -> Int = { it + a }
            val g: (Int) -> Int = { it * a }
            println(f(2))
            println(g(3))
        }
    "#,
    );
    assert_eq!(out, &["3", "3"]);
}

#[test]
fn test_function_type_with_composition() {
    let out = run_prints(
        r#"
        fun comp(a: (Int) -> Int, b: (Int) -> Int): (Int) -> Int {
            return { x -> a(b(x)) }
        }
        fun main() {
            val f = comp({ it + 1 }, { it * 2 })
            println(f(3))
        }
    "#,
    );
    assert_eq!(out, &["7"]);
}

#[test]
fn test_function_type_with_returning_unit_and_side_effect() {
    let out = run_prints(
        r#"
        fun execute(values: List<Int>, op: (Int) -> Unit): String {
            for (v in values) op(v)
            return "done"
        }
        fun main() {
            val r = execute(listOf(1, 2)) { println("v" + it) }
            println(r)
        }
    "#,
    );
    assert_eq!(out, &["v1", "v2", "done"]);
}

#[test]
fn test_function_type_with_null_function_no_call() {
    let out = run_prints(
        r#"
        fun runMaybe(v: Int, fn: ((Int) -> Int)?): Int {
            return fn?.invoke(v) ?: 0
        }
        fun main() {
            println(runMaybe(1, null))
        }
    "#,
    );
    assert_eq!(out, &["0"]);
}

#[test]
fn test_function_type_to_string_not_callable() {
    let out = run_prints(
        r#"
        fun main() {
            val f: (Int) -> Int = { it + 1 }
            println(f is Function1<*, *>)
            println(f::class.simpleName != null)
        }
    "#,
    );
    assert_eq!(out, &["true", "true"]);
}

#[test]
fn test_function_type_through_map_entry() {
    let out = run_prints(
        r#"
        fun main() {
            val map: Map<String, (Int) -> Int> = mapOf(
                "a" to { it + 1 },
                "b" to { it * 2 }
            )
            println(map["a"]?.invoke(3))
            println(map["b"]?.invoke(3))
        }
    "#,
    );
    assert_eq!(out, &["4", "6"]);
}

#[test]
fn test_function_type_receiver_lambda_call_operator() {
    let out = run_prints(
        r#"
        fun applyBlock(v: Int, op: Int.() -> Int): Int {
            return v.op()
        }
        fun main() {
            // Was `fun Int.() -> Int { … }` — not Kotlin: an anonymous
            // function spells its return type with a COLON.
            val inc = fun Int.(): Int { return this + 1 }
            println(applyBlock(5, inc))
        }
    "#,
    );
    assert_eq!(out, &["6"]);
}

#[test]
fn test_function_type_in_array_transform() {
    let out = run_prints(
        r#"
        fun main() {
            val ops: Array<(Int) -> Int> = arrayOf({ it + 1 }, { it - 1 })
            var value = 10
            for (op in ops) {
                value = op(value)
            }
            println(value)
        }
    "#,
    );
    assert_eq!(out, &["10"]);
}

#[test]
fn test_function_type_noarg_constructor() {
    let out = run_prints(
        r#"
        val defaultFactory: () -> String = { "ok" }
        fun main() {
            println(defaultFactory())
        }
    "#,
    );
    assert_eq!(out, &["ok"]);
}

#[test]
fn test_function_type_pass_through_higher_order() {
    let out = run_prints(
        r#"
        fun wrap(f: (Int) -> Int): (Int) -> Int = { n -> f(n) + 1 }
        fun main() {
            val base: (Int) -> Int = { it * 2 }
            val wrapped = wrap(base)
            println(wrapped(3))
        }
    "#,
    );
    assert_eq!(out, &["7"]);
}

#[test]
fn test_function_type_in_constructor() {
    let out = run_prints(
        r#"
        class Processor(val op: (Int) -> Int) {
            fun run(v: Int): Int = op(v)
        }
        fun main() {
            println(Processor { it + 4 }.run(2))
        }
    "#,
    );
    assert_eq!(out, &["6"]);
}

#[test]
fn test_function_type_with_data_class_property() {
    let out = run_prints(
        r#"
        data class Worker(val transform: (Int) -> Int)
        fun main() {
            val w = Worker { it * it }
            println(w.transform(5))
        }
    "#,
    );
    assert_eq!(out, &["25"]);
}

#[test]
fn test_function_type_in_try_catch_flow() {
    let out = run_prints(
        r#"
        fun dispatch(v: Int, fn: (Int) -> Int): Int {
            return if (v < 0) 0 else fn(v)
        }
        fun main() {
            println(dispatch(3, { it + 1 }))
            println(dispatch(-2, { it + 1 }))
        }
    "#,
    );
    assert_eq!(out, &["4", "0"]);
}
