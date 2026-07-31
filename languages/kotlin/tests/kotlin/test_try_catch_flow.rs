kotlin_run_test!(
    test_try_catch_basic,
    r#"
        fun main() {
            try {
                throw Exception("boom")
            } catch (e: Exception) {
                println("caught")
            }
        }
    "#,
    &["caught"]
);

kotlin_run_test!(
    test_try_no_exception,
    r#"
        fun main() {
            try {
                println("ok")
            } catch (e: Exception) {
                println("fail")
            }
        }
    "#,
    &["ok"]
);

kotlin_run_test!(
    test_try_finally_always_runs,
    r#"
        fun main() {
            try {
                println("try")
            } finally {
                println("finally")
            }
        }
    "#,
    &["try", "finally"]
);

kotlin_run_test!(
    test_try_catch_finally_order,
    r#"
        fun main() {
            try {
                throw Exception("x")
            } catch (e: Exception) {
                println("catch")
            } finally {
                println("finally")
            }
            println("after")
        }
    "#,
    &["catch", "finally", "after"]
);

kotlin_run_test!(
    test_try_multiple_catch_branches,
    r#"
        fun main() {
            try {
                throw Exception("x")
            } catch (e: IllegalArgumentException) {
                println("illegal")
            } catch (e: Exception) {
                println("general")
            }
        }
    "#,
    &["general"]
);

kotlin_run_test!(
    test_try_catch_then_return,
    r#"
        fun compute(x: Int): Int {
            try {
                if (x < 0) throw Exception("bad")
                return x + 1
            } catch (e: Exception) {
                return -1
            }
        }
        fun main() {
            println(compute(1))
            println(compute(-1))
        }
    "#,
    &["2", "-1"]
);

kotlin_run_test!(
    test_try_nested_catch,
    r#"
        fun main() {
            try {
                try {
                    throw Exception("inner")
                } catch (e: Exception) {
                    println("inner")
                    throw e
                }
            } catch (e: Exception) {
                println("outer")
            }
        }
    "#,
    &["inner", "outer"]
);

kotlin_run_test!(
    test_catch_only_finally,
    r#"
        fun main() {
            try {
                throw Exception("x")
            } finally {
                println("ok")
            }
        }
    "#,
    &["ok"]
);

kotlin_run_test!(
    test_nested_try_no_throw,
    r#"
        fun main() {
            try {
                val ok = try {
                    2 + 2
                } catch (e: Exception) {
                    -1
                }
                println(ok)
            } catch (e: Exception) {
                println("outer")
            }
        }
    "#,
    &["4"]
);

kotlin_run_test!(
    test_try_with_boolean_guard,
    r#"
        fun guard(x: Int): Boolean = x > 0
        fun main() {
            val x = try {
                if (guard(1)) 10 else throw Exception("bad")
            } catch (e: Exception) {
                0
            }
            println(x)
        }
    "#,
    &["10"]
);

kotlin_run_test!(
    test_try_shadowed_variable,
    r#"
        fun main() {
            val x = 1
            try {
                val x = 5
                println(x)
            } catch (e: Exception) {
                println(0)
            }
            println(x)
        }
    "#,
    &["5", "1"]
);

kotlin_run_test!(
    test_try_catch_in_loop,
    r#"
        fun maybe(i: Int): Int {
            if (i < 0) throw Exception("neg")
            return i
        }
        fun main() {
            var sum = 0
            for (i in -1..2) {
                try {
                    sum += maybe(i)
                } catch (e: Exception) {
                    sum += 10
                }
            }
            println(sum)
        }
    "#,
    &["12"]
);

kotlin_run_test!(
    test_try_rethrow_in_catch,
    r#"
        fun main() {
            try {
                try {
                    throw Exception("x")
                } catch (e: Exception) {
                    throw RuntimeException("wrapped")
                }
            } catch (e: RuntimeException) {
                println("wrapped")
            }
        }
    "#,
    &["wrapped"]
);

kotlin_run_test!(
    test_try_finally_without_exception_result,
    r#"
        fun main() {
            val x = try {
                3 * 4
            } finally {
                println("cleanup")
            }
            println(x)
        }
    "#,
    &["cleanup", "12"]
);

kotlin_run_test!(
    test_try_finally_with_exception,
    r#"
        fun main() {
            try {
                val x = 1 / 0
                println(x)
            } finally {
                println("finally")
            }
        }
    "#,
    &["finally"]
);

kotlin_run_test!(
    test_try_catch_return_value_preserved,
    r#"
        fun safeDivide(a: Int, b: Int): Int {
            return try {
                a / b
            } catch (e: Exception) {
                0
            }
        }
        fun main() {
            println(safeDivide(10, 2))
            println(safeDivide(10, 0))
        }
    "#,
    &["5", "0"]
);

kotlin_run_test!(
    test_try_with_custom_message,
    r#"
        fun main() {
            try {
                throw IllegalArgumentException("bad")
            } catch (e: IllegalArgumentException) {
                println(e.message)
            }
        }
    "#,
    &["bad"]
);

kotlin_run_test!(
    test_try_with_nested_finally,
    r#"
        fun main() {
            try {
                try {
                    println("inner")
                } finally {
                    println("inner-finally")
                }
            } finally {
                println("outer-finally")
            }
        }
    "#,
    &["inner", "inner-finally", "outer-finally"]
);

kotlin_run_test!(
    test_try_catch_finally_expression_assignment,
    r#"
        fun main() {
            val value = try {
                throw Exception("x")
            } catch (e: Exception) {
                8
            } finally {
                println("cleanup")
            }
            println(value)
        }
    "#,
    &["cleanup", "8"]
);

kotlin_run_test!(
    test_try_chain_and_return,
    r#"
        fun branch(x: Int): Int {
            return try {
                if (x == 0) throw Exception("0")
                x * 2
            } catch (e: Exception) {
                -1
            }
        }
        fun main() {
            println(branch(3))
            println(branch(0))
        }
    "#,
    &["6", "-1"]
);

kotlin_run_test!(
    test_throw_in_loop_continue,
    r#"
        fun main() {
            var out = 0
            for (i in 0..3) {
                try {
                    if (i == 2) throw Exception("stop")
                    out += i
                } catch (e: Exception) {
                    out += 10
                }
            }
            println(out)
        }
    "#,
    &["12"]
);

kotlin_run_test!(
    test_try_with_boolean_catch,
    r#"
        fun main() {
            val ok = try {
                true
            } catch (e: Exception) {
                false
            }
            println(ok)
        }
    "#,
    &["true"]
);

kotlin_run_test!(
    test_try_catch_with_custom_error,
    r#"
        class AppError(msg: String) : Exception(msg)
        fun main() {
            try {
                throw AppError("boom")
            } catch (e: AppError) {
                println(e.message)
            }
        }
    "#,
    &["boom"]
);

kotlin_run_test!(
    test_try_finally_and_return,
    r#"
        fun value(x: Int): Int {
            try {
                return x * 2
            } finally {
                println("done")
            }
        }
        fun main() {
            println(value(3))
        }
    "#,
    &["done", "6"]
);

kotlin_run_test!(
    test_try_nested_multi_catch,
    r#"
        fun main() {
            try {
                throw IllegalArgumentException("bad")
            } catch (e: RuntimeException) {
                println("runtime")
            } catch (e: Exception) {
                println("general")
            }
        }
    "#,
    &["runtime"]
);

kotlin_run_test!(
    test_try_finally_with_try_result,
    r#"
        fun main() {
            val x = try {
                100
            } finally {
                println("f")
            }
            println(x)
        }
    "#,
    &["f", "100"]
);

kotlin_run_test!(
    test_try_catch_in_expression_chain,
    r#"
        fun value(x: Int): Int {
            return try {
                if (x == 0) throw IllegalStateException()
                10 / x
            } catch (e: IllegalStateException) {
                -1
            } catch (e: Exception) {
                -2
            }
        }
        fun main() {
            println(value(0))
            println(value(2))
        }
    "#,
    &["-1", "5"]
);

kotlin_run_test!(
    test_try_with_throwed_boolean,
    r#"
        fun main() {
            try {
                if (true) throw Exception("x")
                println("never")
            } catch (e: Exception) {
                println(e.message)
            } finally {
                println("done")
            }
        }
    "#,
    &["x", "done"]
);

kotlin_run_test!(
    test_try_resource_like_sequence,
    r#"
        fun main() {
            var opened = 0
            try {
                opened += 1
                try {
                    opened += 10
                } finally {
                    opened += 100
                }
            } finally {
                opened += 1000
            }
            println(opened)
        }
    "#,
    &["1111"]
);

kotlin_run_test!(
    test_try_nested_loop_and_error,
    r#"
        fun main() {
            var out = 0
            for (i in 1..3) {
                try {
                    if (i == 2) throw Exception("x")
                    out += i
                } catch (e: Exception) {
                    out += 10
                }
            }
            println(out)
        }
    "#,
    &["12"]
);
