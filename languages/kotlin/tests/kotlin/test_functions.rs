use crate::helpers::run_prints;

#[test]
fn test_function_expression_body() {
    let out = run_prints(
        r#"
        fun square(x: Int): Int = x * x

        fun main() {
            println(square(6))
        }
    "#,
    );
    assert_eq!(out, &["36"]);
}

#[test]
fn test_function_block_body() {
    let out = run_prints(
        r#"
        fun greet(name: String) {
            println("Hi " + name)
        }

        fun main() {
            greet("Kotlin")
        }
    "#,
    );
    assert_eq!(out, &["Hi Kotlin"]);
}

#[test]
fn test_function_recursion() {
    let out = run_prints(
        r#"
        fun fact(n: Int): Int {
            if (n <= 1) {
                return 1
            }
            return n * fact(n - 1)
        }

        fun main() {
            println(fact(5))
        }
    "#,
    );
    assert_eq!(out, &["120"]);
}

#[test]
fn test_named_arguments() {
    let out = run_prints(
        r#"
        fun formatName(first: String, last: String): String {
            return last + ", " + first
        }

        fun main() {
            println(formatName(last = "Smith", first = "Alice"))
        }
    "#,
    );
    assert_eq!(out, &["Smith, Alice"]);
}

#[test]
fn test_function_multiple_parameters() {
    let out = run_prints(
        r#"
        fun add(a: Int, b: Int, c: Int): Int {
            return a + b + c
        }

        fun main() {
            println(add(1, 2, 3))
        }
    "#,
    );
    assert_eq!(out, &["6"]);
}

#[test]
fn test_function_void_return() {
    let out = run_prints(
        r#"
        fun doNothing() {
            val a = 1
        }

        fun main() {
            doNothing()
            println("done")
        }
    "#,
    );
    assert_eq!(out, &["done"]);
}

#[test]
fn test_function_distinct_names() {
    let out = run_prints(
        r#"
        fun printOne(a: Int) {
            println(a)
        }

        fun printTwo(a: Int, b: Int) {
            println(a + b)
        }

        fun main() {
            printOne(10)
            printTwo(10, 20)
        }
    "#,
    );
    assert_eq!(out, &["10", "30"]);
}

#[test]
fn test_function_calling_another_function() {
    let out = run_prints(
        r#"
        fun doubleVal(x: Int): Int = x * 2
        fun tripleVal(x: Int): Int = doubleVal(x) + x

        fun main() {
            println(tripleVal(4))
        }
    "#,
    );
    assert_eq!(out, &["12"]);
}

#[test]
fn test_function_early_return() {
    let out = run_prints(
        r#"
        fun checkPositive(x: Int) {
            if (x <= 0) return
            println("positive")
        }

        fun main() {
            checkPositive(-5)
            checkPositive(5)
        }
    "#,
    );
    assert_eq!(out, &["positive"]);
}

#[test]
fn test_function_boolean_return() {
    let out = run_prints(
        r#"
        fun isEven(n: Int): Boolean = (n % 2 == 0)

        fun main() {
            println(isEven(4))
            println(isEven(7))
        }
    "#,
    );
    assert_eq!(out, &["true", "false"]);
}

#[test]
fn test_function_string_return() {
    let out = run_prints(
        r#"
        fun concat(a: String, b: String): String = a + b

        fun main() {
            println(concat("foo", "bar"))
        }
    "#,
    );
    assert_eq!(out, &["foobar"]);
}

#[test]
fn test_function_parameter_shadowing() {
    let out = run_prints(
        r#"
        val x = 100

        fun testShadow(x: Int) {
            println(x)
        }

        fun main() {
            testShadow(5)
            println(x)
        }
    "#,
    );
    assert_eq!(out, &["5", "100"]);
}

#[test]
fn test_function_chained_calls() {
    let out = run_prints(
        r#"
        fun inc(x: Int): Int = x + 1

        fun main() {
            println(inc(inc(inc(0))))
        }
    "#,
    );
    assert_eq!(out, &["3"]);
}

#[test]
fn test_function_expression_body_logic() {
    let out = run_prints(
        r#"
        fun maxOf(a: Int, b: Int): Int = if (a > b) a else b

        fun main() {
            println(maxOf(10, 20))
        }
    "#,
    );
    assert_eq!(out, &["20"]);
}

#[test]
fn test_function_fibonacci() {
    let out = run_prints(
        r#"
        fun fib(n: Int): Int {
            if (n <= 1) return n
            return fib(n - 1) + fib(n - 2)
        }

        fun main() {
            println(fib(6))
        }
    "#,
    );
    assert_eq!(out, &["8"]);
}

#[test]
fn test_function_multiple_returns() {
    let out = run_prints(
        r#"
        fun sign(x: Int): Int {
            if (x > 0) return 1
            if (x < 0) return -1
            return 0
        }

        fun main() {
            println(sign(10))
            println(sign(-5))
            println(sign(0))
        }
    "#,
    );
    assert_eq!(out, &["1", "-1", "0"]);
}

#[test]
fn test_function_local_var_mutation() {
    let out = run_prints(
        r#"
        fun accum(n: Int): Int {
            var sum = 0
            var i = 1
            while (i <= n) {
                sum += i
                i += 1
            }
            return sum
        }

        fun main() {
            println(accum(4))
        }
    "#,
    );
    assert_eq!(out, &["10"]);
}

#[test]
fn test_function_default_parameter() {
    let out = run_prints(
        r#"
        fun greet(name: String = "friend") {
            println("hi " + name)
        }

        fun main() {
            greet()
            greet("Dev")
        }
    "#,
    );
    assert_eq!(out, &["hi friend", "hi Dev"]);
}

#[test]
fn test_function_local_nested() {
    let out = run_prints(
        r#"
        fun main() {
            fun triple(x: Int): Int {
                return x * 3
            }
            println(triple(4))
        }
    "#,
    );
    assert_eq!(out, &["12"]);
}

#[test]
fn test_function_no_return_uses_unit() {
    let out = run_prints(
        r#"
        fun sideEffect() {
            println("started")
        }

        fun main() {
            sideEffect()
            println("done")
        }
    "#,
    );
    assert_eq!(out, &["started", "done"]);
}

#[test]
fn test_function_optional_else_return() {
    let out = run_prints(
        r#"
        fun maybePositive(n: Int): String {
            if (n > 0) return "positive"
            return "not positive"
        }

        fun main() {
            println(maybePositive(3))
            println(maybePositive(-1))
        }
    "#,
    );
    assert_eq!(out, &["positive", "not positive"]);
}

#[test]
fn test_function_call_chain_with_hof() {
    let out = run_prints(
        r#"
        fun transform(x: Int, op: (Int) -> Int): Int {
            return op(x)
        }

        fun main() {
            val result = transform(10) { it + 5 }
            println(result)
        }
    "#,
    );
    assert_eq!(out, &["15"]);
}

#[test]
fn test_function_varargs() {
    let out = run_prints(
        r#"
        fun sumAll(vararg values: Int): Int {
            var total = 0
            var i = 0
            while (i < 4) {
                total += values[i]
                i += 1
            }
            return total
        }

        fun main() {
            println(sumAll(1, 2, 3, 4))
        }
    "#,
    );
    assert_eq!(out, &["10"]);
}

#[test]
fn test_function_default_and_named_args() {
    let out = run_prints(
        r#"
        fun greet(prefix: String = "Hi", name: String) {
            println(prefix + " " + name)
        }

        fun main() {
            greet(name = "Kotlin")
            greet(prefix = "Hello", name = "Rust")
        }
    "#,
    );
    assert_eq!(out, &["Hi Kotlin", "Hello Rust"]);
}

#[test]
fn test_function_pass_function_as_value() {
    let out = run_prints(
        r#"
        fun apply(value: Int, op: (Int) -> Int): Int {
            return op(value)
        }

        fun main() {
            val transform = { x: Int -> x * 3 }
            println(apply(4, transform))
        }
    "#,
    );
    assert_eq!(out, &["12"]);
}

#[test]
fn test_function_local_nested_calls() {
    let out = run_prints(
        r#"
        fun main() {
            fun base(x: Int): Int { return x * 2 }
            fun nested(x: Int): Int { return base(x) + 1 }
            println(nested(3))
        }
    "#,
    );
    assert_eq!(out, &["7"]);
}

#[test]
fn test_function_returning_lambda() {
    let out = run_prints(
        r#"
        fun makeMultiplier(mult: Int): (Int) -> Int {
            return { value -> value * mult }
        }

        fun main() {
            val timesFive = makeMultiplier(5)
            println(timesFive(3))
        }
    "#,
    );
    assert_eq!(out, &["15"]);
}

#[test]
fn test_function_with_no_args_named_call() {
    let out = run_prints(
        r#"
        fun getStatus(prefix: String = "OK", code: Int = 200): String {
            return prefix + code.toString()
        }

        fun main() {
            println(getStatus())
            println(getStatus(code = 404))
        }
    "#,
    );
    assert_eq!(out, &["OK200", "OK404"]);
}

#[test]
fn test_function_boolean_contract() {
    let out = run_prints(
        r#"
        fun allPositive(a: Int, b: Int): Boolean {
            return a > 0 && b > 0
        }

        fun main() {
            println(allPositive(3, 4))
            println(allPositive(3, -1))
        }
    "#,
    );
    assert_eq!(out, &["true", "false"]);
}

#[test]
fn test_function_while_loop_in_function() {
    let out = run_prints(
        r#"
        fun sumUpTo(n: Int): Int {
            var i = 1
            var total = 0
            while (i <= n) {
                total += i
                i += 1
            }
            return total
        }

        fun main() {
            println(sumUpTo(5))
        }
    "#,
    );
    assert_eq!(out, &["15"]);
}

#[test]
fn test_function_overload_resolution_by_type() {
    let out = run_prints(
        r#"
        fun format(value: Int): String = "int:" + value
        fun format(value: String): String = "str:" + value

        fun main() {
            println(format(7))
            println(format("7"))
        }
    "#,
    );
    assert_eq!(out, &["int:7", "str:7"]);
}

#[test]
fn test_function_uses_tailrec_optimization_contract() {
    let out = run_prints(
        r#"
        fun countdown(start: Int): Int {
            tailrec fun loop(current: Int, acc: Int): Int {
                if (current == 0) return acc
                return loop(current - 1, acc + current)
            }
            return loop(start, 0)
        }

        fun main() {
            println(countdown(4))
            println(countdown(0))
        }
    "#,
    );
    assert_eq!(out, &["10", "0"]);
}

#[test]
fn test_function_vararg_with_spread_operator() {
    let out = run_prints(
        r#"
        fun sum(label: String, vararg values: Int): String {
            var total = 0
            for (value in values) {
                total += value
            }
            return label + ":" + total
        }

        fun main() {
            val extras = intArrayOf(4, 5)
            println(sum("base", 1, 2, 3, *extras))
            println(sum("empty"))
        }
    "#,
    // 1 + 2 + 3 + *[4, 5] sums to 15 (real Kotlin agrees).
    );
    assert_eq!(out, &["base:15", "empty:0"]);
}

#[test]
fn test_function_reference_invocation() {
    let out = run_prints(
        r#"
        fun double(x: Int): Int = x * 2

        fun main() {
            val f: (Int) -> Int = ::double
            println(f(7))
            println(::double(3))
        }
    "#,
    );
    assert_eq!(out, &["14", "6"]);
}

#[test]
fn test_function_extension_member() {
    let out = run_prints(
        r#"
        fun String.wrap(prefix: String): String {
            return prefix + this
        }

        fun main() {
            println("kotlin".wrap("["))
            println("v".wrap("v"))
        }
    "#,
    );
    assert_eq!(out, &["[kotlin", "vv"]);
}

#[test]
fn test_function_nullable_parameter_returns_default() {
    let out = run_prints(
        r#"
        fun describe(input: String?): Int {
            return if (input == null) 0 else input.length
        }

        fun main() {
            println(describe(null))
            println(describe("abc"))
        }
    "#,
    );
    assert_eq!(out, &["0", "3"]);
}

#[test]
fn test_function_generic_identity_and_inference() {
    let out = run_prints(
        r#"
        fun <T> identity(value: T): T = value

        fun main() {
            println(identity("text"))
            println(identity(12))
            println(identity(true))
        }
    "#,
    );
    assert_eq!(out, &["text", "12", "true"]);
}

#[test]
fn test_function_return_type_inference() {
    let out = run_prints(
        r#"
        fun compute(value: Int) = if (value > 0) value.toString() else value

        fun main() {
            println(compute(4))
            println(compute(-1))
        }
    "#,
    );
    assert_eq!(out, &["4", "-1"]);
}

#[test]
fn test_function_default_and_named_mix() {
    let out = run_prints(
        r#"
        fun build(prefix: String = "p", suffix: String = "s", count: Int = 1): String {
            return prefix + suffix + count.toString()
        }

        fun main() {
            println(build())
            println(build(count = 3))
            println(build("a", count = 4, suffix = "b"))
        }
    "#,
    );
    assert_eq!(out, &["ps1", "ps3", "ab4"]);
}

#[test]
fn test_function_nested_capture_uses_outer_state() {
    let out = run_prints(
        r#"
        fun makeAdder(base: Int): (Int) -> Int {
            return { delta ->
                base + delta
            }
        }

        fun main() {
            val addTen = makeAdder(10)
            println(addTen(3))
        }
    "#,
    );
    assert_eq!(out, &["13"]);
}

#[test]
fn test_function_default_arg_from_previous_parameter() {
    let out = run_prints(
        r#"
        fun tag(base: String, suffix: String = base.uppercase()): String {
            return base + ":" + suffix
        }

        fun main() {
            println(tag("kotlin"))
            println(tag("kotlin", "custom"))
        }
    "#,
    );
    assert_eq!(out, &["kotlin:KOTLIN", "kotlin:custom"]);
}

#[test]
fn test_function_higher_order_with_default_lambda() {
    let out = run_prints(
        r#"
        fun apply(value: Int, op: (Int) -> Int = { it + 1 }): Int {
            return op(value)
        }

        fun main() {
            println(apply(4))
            println(apply(4) { it * 3 })
        }
    "#,
    );
    assert_eq!(out, &["5", "12"]);
}

#[test]
fn test_function_local_tailrec_accumulator() {
    let out = run_prints(
        r#"
        fun power(base: Int, exp: Int): Int {
            tailrec fun loop(remaining: Int, acc: Int): Int {
                if (remaining == 0) return acc
                return loop(remaining - 1, acc * base)
            }
            return loop(exp, 1)
        }

        fun main() {
            println(power(2, 0))
            println(power(3, 3))
        }
    "#,
    );
    assert_eq!(out, &["1", "27"]);
}

#[test]
fn test_function_mutually_recursive_parity() {
    let out = run_prints(
        r#"
        fun isEven(n: Int): Boolean {
            if (n == 0) return true
            return isOdd(n - 1)
        }

        fun isOdd(n: Int): Boolean {
            if (n == 0) return false
            return isEven(n - 1)
        }

        fun main() {
            println(isEven(8))
            println(isOdd(8))
            println(isEven(7))
        }
    "#,
    );
    assert_eq!(out, &["true", "false", "false"]);
}

#[test]
fn test_function_throws_and_finally_blocks_do_not_swallow_return_value() {
    let out = run_prints(
        r#"
        fun risky(value: Int): Int {
            try {
                if (value < 0) {
                    throw Exception("bad")
                }
                return value * 2
            } finally {
                println("final")
            }
        }

        fun main() {
            println(risky(4))
            try {
                risky(-1)
            } catch (e: Exception) {
                println("caught")
            }
        }
    "#,
    );
    assert_eq!(out, &["final", "8", "final", "caught"]);
}
