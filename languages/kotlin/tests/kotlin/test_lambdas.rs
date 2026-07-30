use crate::helpers::run_prints;

#[test]
fn test_lambda_expression() {
    let out = run_prints(r#"
        fun main() {
            val mult = { a: Int, b: Int -> a * b }
            println(mult(4, 5))
        }
    "#);
    assert_eq!(out, &["20"]);
}

#[test]
fn test_implicit_it_lambda() {
    let out = run_prints(r#"
        fun main() {
            val doubleIt = { it * 2 }
            println(doubleIt(21))
        }
    "#);
    assert_eq!(out, &["42"]);
}

#[test]
fn test_higher_order_function() {
    let out = run_prints(r#"
        fun execute(a: Int, b: Int, op: (Int, Int) -> Int): Int {
            return op(a, b)
        }

        fun main() {
            val sum = execute(15, 25, { x, y -> x + y })
            println(sum)
        }
    "#);
    assert_eq!(out, &["40"]);
}

#[test]
fn test_trailing_lambda_syntax() {
    let out = run_prints(r#"
        fun applyOp(x: Int, op: (Int) -> Int): Int {
            return op(x)
        }

        fun main() {
            val result = applyOp(10) { it * 3 }
            println(result)
        }
    "#);
    assert_eq!(out, &["30"]);
}

#[test]
fn test_no_arg_lambda() {
    let out = run_prints(r#"
        fun main() {
            val sayHi = { "Hi" }
            println(sayHi())
        }
    "#);
    assert_eq!(out, &["Hi"]);
}

#[test]
fn test_lambda_closure_read() {
    let out = run_prints(r#"
        fun main() {
            val factor = 10
            val mult = { x: Int -> x * factor }
            println(mult(5))
        }
    "#);
    assert_eq!(out, &["50"]);
}

#[test]
fn test_lambda_in_variable() {
    let out = run_prints(r#"
        fun main() {
            val sub = { a: Int, b: Int -> a - b }
            println(sub(100, 30))
        }
    "#);
    assert_eq!(out, &["70"]);
}

#[test]
fn test_lambda_returning_boolean() {
    let out = run_prints(r#"
        fun main() {
            val isGreater = { a: Int, b: Int -> a > b }
            println(isGreater(10, 5))
            println(isGreater(2, 8))
        }
    "#);
    assert_eq!(out, &["true", "false"]);
}

#[test]
fn test_lambda_with_no_arguments() {
    let out = run_prints(r#"
        fun main() {
            val supplier = { "value" }
            println(supplier())
        }
    "#);
    assert_eq!(out, &["value"]);
}

#[test]
fn test_lambda_returning_function() {
    let out = run_prints(r#"
        fun makeAdder(offset: Int): (Int) -> Int {
            return { value -> value + offset }
        }

        fun main() {
            val addFive = makeAdder(5)
            println(addFive(10))
        }
    "#);
    assert_eq!(out, &["15"]);
}

#[test]
fn test_lambda_with_branching() {
    let out = run_prints(r#"
        fun main() {
            val check = { value: Int ->
                if (value > 10) {
                    "big"
                } else {
                    "small"
                }
            }
            println(check(3))
            println(check(15))
        }
    "#);
    assert_eq!(out, &["small", "big"]);
}

#[test]
fn test_lambda_as_argument_with_trailing_syntax() {
    let out = run_prints(r#"
        fun applyTwice(value: Int, op: (Int) -> Int): Int {
            return op(op(value))
        }

        fun main() {
            val result = applyTwice(3) { it * 2 }
            println(result)
        }
    "#);
    assert_eq!(out, &["12"]);
}

#[test]
fn test_lambda_with_three_parameters() {
    let out = run_prints(r#"
        fun main() {
            val combine = { a: Int, b: Int, c: Int -> a + b + c }
            println(combine(1, 2, 3))
        }
    "#);
    assert_eq!(out, &["6"]);
}

#[test]
fn test_lambda_returning_lambda() {
    let out = run_prints(r#"
        fun makeMultiplier(scale: Int): (Int) -> Int {
            return { x: Int -> x * scale }
        }

        fun main() {
            val times = makeMultiplier(4)
            val plus = makeMultiplier(1)
            println(times(2))
            println(plus(3))
        }
    "#);
    assert_eq!(out, &["8", "3"]);
}

#[test]
fn test_lambda_with_if_expression() {
    let out = run_prints(r#"
        fun main() {
            val evaluator = { n: Int ->
                if (n > 10) {
                    "big"
                } else {
                    "small"
                }
            }
            println(evaluator(2))
            println(evaluator(20))
        }
    "#);
    assert_eq!(out, &["small", "big"]);
}

#[test]
fn test_lambda_and_local_mutation() {
    let out = run_prints(r#"
        fun main() {
            var total = 0
            val add = { value: Int ->
                total += value
            }
            add(3)
            add(5)
            println(total)
        }
    "#);
    assert_eq!(out, &["8"]);
}

#[test]
fn test_lambda_as_default_argument() {
    let out = run_prints(r#"
        fun calculate(input: Int, op: (Int) -> Int = { it * 2 }): Int {
            return op(input)
        }

        fun main() {
            println(calculate(4))
            println(calculate(4, { it + 1 }))
        }
    "#);
    assert_eq!(out, &["8", "5"]);
}

#[test]
fn test_lambda_nested_call() {
    let out = run_prints(r#"
        fun make(): (Int) -> Int {
            return { value -> value + 1 }
        }

        fun main() {
            val fnRef = make()
            println(fnRef(6))
            println(make()(3))
        }
    "#);
    assert_eq!(out, &["7", "4"]);
}

#[test]
fn test_lambda_with_block_body() {
    let out = run_prints(r#"
        fun main() {
            val compute = { value: Int ->
                val x = value * 2
                val y = x + 1
                y
            }
            println(compute(5))
        }
    "#);
    assert_eq!(out, &["11"]);
}

#[test]
fn test_lambda_as_expression_argument_without_params() {
    let out = run_prints(r#"
        fun runTask(task: () -> String): String {
            return task()
        }

        fun main() {
            println(runTask({ "done" }))
        }
    "#);
    assert_eq!(out, &["done"]);
}

#[test]
fn test_lambda_takes_lambda() {
    let out = run_prints(r#"
        fun withOperation(value: Int, transform: (Int) -> Int): Int {
            return transform(value)
        }

        fun main() {
            val pipeline = withOperation
            println(pipeline(5, { it + 10 }))
            println(pipeline(2, { it * 3 }))
        }
    "#);
    assert_eq!(out, &["15", "6"]);
}

#[test]
fn test_lambda_as_argument_list_value() {
    let out = run_prints(r#"
        fun apply(x: Int, y: Int, op: (Int, Int) -> Int): Int {
            return op(x, y)
        }

        fun main() {
            println(apply(8, 4, { a, b -> a - b }))
            println(apply(2, 2, { a, b -> a / b }))
        }
    "#);
    assert_eq!(out, &["4", "1"]);
}

#[test]
fn test_lambda_returning_unit() {
    let out = run_prints(r#"
fun main() { val action: (String) -> Unit = { s -> println(s) }; action("go") }
"#);
    assert_eq!(out, &["go"]);
}

#[test]
fn test_lambda_capture_and_modify() {
    let out = run_prints(r#"
fun main() { var base = 1; val inc = { x: Int -> base + x }; println(inc(4)) }
"#);
    assert_eq!(out, &["5"]);
}

#[test]
fn test_lambda_two_returns() {
    let out = run_prints(r#"
fun main() { val choose = { flag: Boolean -> if (flag) "yes" else "no" }; println(choose(true)); println(choose(false)) }
"#);
    assert_eq!(out, &["yes", "no"]);
}

#[test]
fn test_lambda_as_map_operation() {
    let out = run_prints(r#"
fun transform(v: Int, op: (Int) -> Int): Int { return op(v) }; fun main() { println(transform(3, { it + 2 })); println(transform(4, { x -> x * 2 })) }
"#);
    assert_eq!(out, &["5", "8"]);
}

#[test]
fn test_lambda_parameter_default_value() {
    let out = run_prints(r#"
fun main() { val add = { base: Int, bonus: Int -> base + bonus }; println(add(6, 4)) }
"#);
    assert_eq!(out, &["10"]);
}

#[test]
fn test_lambda_in_nested_function() {
    let out = run_prints(r#"
fun factory(base: Int): (Int) -> Int { return { v -> v + base } }; fun main() { val addTen = factory(10); println(addTen(5)) }
"#);
    assert_eq!(out, &["15"]);
}

#[test]
fn test_lambda_with_multiple_calls() {
    let out = run_prints(r#"
fun runTwice(v: Int, op: (Int) -> Int): Int { return op(op(v)) }; fun main() { println(runTwice(2, { it * 3 })) }
"#);
    assert_eq!(out, &["18"]);
}

#[test]
fn test_lambda_boolean_check() {
    let out = run_prints(r#"
fun main() { val isEven = { x: Int -> x % 2 == 0 }; println(isEven(8)); println(isEven(5)) }
"#);
    assert_eq!(out, &["true", "false"]);
}

#[test]
fn test_lambda_stored_in_array() {
    let out = run_prints(r#"
fun main() { val ops = arrayOf({ x: Int -> x + 1 }, { x: Int -> x * 2 }); println(ops[0](3)); println(ops[1](3)) }
"#);
    assert_eq!(out, &["4", "6"]);
}

#[test]
fn test_lambda_with_receiver_style_extension_call() {
    let out = run_prints(r#"
fun main() {
    val shout: String.() -> String = { this.uppercase() + "!" }
    println("go".shout())
}
"#);
    assert_eq!(out, &["GO!"]);
}

#[test]
fn test_lambda_with_destructured_pair_param() {
    let out = run_prints(r#"
fun transform(input: Pair<Int, Int>, op: (Int, Int) -> Int): Int {
    return op(input.first, input.second)
}

fun main() {
    println(transform(Pair(3, 4), { a, b -> a + b }))
    println(transform(Pair(10, 2), { a, b -> a - b }))
}
"#);
    assert_eq!(out, &["7", "8"]);
}

#[test]
fn test_lambda_nullable_return_contract() {
    let out = run_prints(r#"
fun resolve(flag: Boolean): (Int) -> Int? {
    return if (flag) {
        { x -> x + 1 }
    } else {
        { _ -> null }
    }
}

fun main() {
    val f = resolve(true)
    val g = resolve(false)
    println(f(1) ?: -1)
    println(g(1) ?: -1)
}
"#);
    assert_eq!(out, &["2", "-1"]);
}

#[test]
fn test_lambda_try_catch_with_local_state() {
    let out = run_prints(r#"
fun main() {
    val safe = { s: String ->
        try {
            s.toInt()
        } catch (e: Exception) {
            -1
        }
    }
    println(safe("12"))
    println(safe("bad"))
}
"#);
    assert_eq!(out, &["12", "-1"]);
}

#[test]
fn test_lambda_with_multiple_parameters_destructuring_syntax() {
    let out = run_prints(r#"
fun main() {
    val merge = { (left, right): Pair<Int, Int> ->
        left + right
    }
    println(merge(Pair(7, 8)))
}
"#);
    assert_eq!(out, &["15"]);
}

#[test]
fn test_lambda_reference_reuse_and_reassignment() {
    let out = run_prints(r#"
fun main() {
    var op: (Int) -> Int = { it + 1 }
    println(op(4))
    op = { it * 2 }
    println(op(4))
}
"#);
    assert_eq!(out, &["5", "8"]);
}

#[test]
fn test_lambda_parameter_named_inference_and_type_annotation() {
    let out = run_prints(r#"
fun <T> apply(values: List<T>, transform: (T) -> Int): Int {
    return values.map(transform).fold(0) { left, right -> left + right }
}

fun main() {
    println(apply(listOf("aa", "b", "ccc"), { it.length }))
}
"#);
    assert_eq!(out, &["6"]);
}
