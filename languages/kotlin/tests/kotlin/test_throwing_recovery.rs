kotlin_run_test!(
    test_throw_and_catch_runtime,
    r#"
        fun fail() = throw IllegalStateException("bad")
        fun main() {
            try {
                fail()
            } catch (e: IllegalStateException) {
                println(e.message)
            }
        }
    "#,
    &["bad"]
);

kotlin_run_test!(
    test_catch_uses_local_state,
    r#"
        var seen = 0
        fun main() {
            try {
                throw Exception("x")
            } catch (e: Exception) {
                seen = 1
            }
            println(seen)
        }
    "#,
    &["1"]
);

kotlin_run_test!(
    test_throwed_type_mismatch,
    r#"
        fun main() {
            try {
                val value: Any = "text"
                val num = value as Int
                println(num)
            } catch (e: ClassCastException) {
                println("cast")
            }
        }
    "#,
    &["cast"]
);

kotlin_run_test!(
    test_throwed_cast_question_mark,
    r#"
        fun main() {
            val value: Any = "text"
            println(value as? Int)
            println(value is Int)
        }
    "#,
    &["null", "false"]
);

kotlin_run_test!(
    test_throwing_custom_class,
    r#"
        class DomainError(message: String) : Exception(message)
        fun main() {
            try {
                throw DomainError("oops")
            } catch (e: DomainError) {
                println(e.message)
            }
        }
    "#,
    &["oops"]
);

kotlin_run_test!(
    test_throwing_string_index_error,
    r#"
        fun main() {
            try {
                val text = "abc"
                println(text[99])
            } catch (e: Exception) {
                println("error")
            }
        }
    "#,
    &["error"]
);

kotlin_run_test!(
    test_throwing_divide_by_zero_recovery,
    r#"
        fun safeDivide(a: Int, b: Int): Int {
            try {
                return a / b
            } catch (e: Exception) {
                return -1
            }
        }
        fun main() {
            println(safeDivide(5, 0))
        }
    "#,
    &["-1"]
);

kotlin_run_test!(
    test_throwing_nested_recover,
    r#"
        fun parseIntOrThrow(s: String): Int {
            if (s == "x") throw NumberFormatException("bad")
            return 1
        }
        fun main() {
            try {
                parseIntOrThrow("x")
            } catch (e: NumberFormatException) {
                println("bad")
            }
        }
    "#,
    &["bad"]
);

kotlin_run_test!(
    test_throwing_multiple_errors,
    r#"
        fun explode(i: Int) {
            when (i) {
                0 -> throw IllegalArgumentException("zero")
                1 -> throw IllegalStateException("state")
                else -> println("ok")
            }
        }
        fun main() {
            try {
                explode(0)
            } catch (e: Exception) {
                println(e::class.java.simpleName)
            }
        }
    "#,
    &["IllegalArgumentException"]
);

kotlin_run_test!(
    test_catch_ordering_specificity,
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
    test_throw_and_finalize,
    r#"
        class Holder {
            fun close() = println("closed")
        }
        fun main() {
            val h = Holder()
            try {
                throw Exception("x")
            } finally {
                h.close()
            }
        }
    "#,
    &["closed"]
);

kotlin_run_test!(
    test_throwing_in_expression_assignment,
    r#"
        fun main() {
            val x = try {
                throw Exception("no")
            } catch (e: Exception) {
                11
            }
            println(x)
        }
    "#,
    &["11"]
);

kotlin_run_test!(
    test_rethrow_new_exception,
    r#"
        fun main() {
            try {
                try {
                    throw Exception("a")
                } catch (e: Exception) {
                    throw RuntimeException("b")
                }
            } catch (e: RuntimeException) {
                println(e.message)
            }
        }
    "#,
    &["b"]
);

kotlin_run_test!(
    test_throwing_and_return_in_catch,
    r#"
        fun check(x: Int): Int {
            try {
                if (x == 0) throw Exception("zero")
                return x
            } catch (e: Exception) {
                return 99
            }
        }
        fun main() {
            println(check(1))
            println(check(0))
        }
    "#,
    &["1", "99"]
);

kotlin_run_test!(
    test_recover_with_default,
    r#"
        fun fallback(x: String): Int {
            try {
                return x.toInt()
            } catch (e: NumberFormatException) {
                return 0
            }
        }
        fun main() {
            println(fallback("12"))
            println(fallback("xx"))
        }
    "#,
    &["12", "0"]
);

kotlin_run_test!(
    test_throwing_for_loop_continue,
    r#"
        fun f(i: Int): Int {
            if (i == 2) throw Exception("bad")
            return i
        }
        fun main() {
            var out = 0
            for (i in 0..3) {
                try {
                    out += f(i)
                } catch (e: Exception) {
                    out += 100
                }
            }
            println(out)
        }
    "#,
    &["103"]
);

kotlin_run_test!(
    test_throwable_chain_and_recover,
    r#"
        fun main() {
            try {
                throw Exception("x")
            } catch (e: Exception) {
                try {
                    println("inner")
                    throw RuntimeException("y")
                } catch (inner: RuntimeException) {
                    println(inner.message)
                }
            }
        }
    "#,
    &["inner", "y"]
);

kotlin_run_test!(
    test_throwing_with_finally_side_effect,
    r#"
        fun main() {
            try {
                throw Exception("stop")
            } finally {
                println("teardown")
            }
        }
    "#,
    &["teardown"]
);

kotlin_run_test!(
    test_try_catch_finally_nested,
    r#"
        fun main() {
            try {
                try {
                    throw Exception("inner")
                } catch (e: Exception) {
                    println("inner")
                } finally {
                    println("inner finally")
                }
            } finally {
                println("outer finally")
            }
        }
    "#,
    &["inner", "inner finally", "outer finally"]
);

kotlin_run_test!(
    test_no_exception_path_in_catch_test,
    r#"
        fun mayFail(shouldFail: Boolean): Int {
            return try {
                if (shouldFail) throw Exception("x")
                1
            } catch (e: Exception) {
                0
            }
        }
        fun main() {
            println(mayFail(false))
            println(mayFail(true))
        }
    "#,
    &["1", "0"]
);

kotlin_run_test!(
    test_throwing_class_cast,
    r#"
        fun main() {
            val x: Any = 10
            try {
                val y = x as String
                println(y)
            } catch (e: ClassCastException) {
                println("class-cast")
            }
        }
    "#,
    &["class-cast"]
);

kotlin_run_test!(
    test_throwing_in_inline_lambda,
    r#"
        fun main() {
            try {
                run {
                    throw Exception("inline")
                }
            } catch (e: Exception) {
                println(e.message)
            }
        }
    "#,
    &["inline"]
);

kotlin_run_test!(
    test_throwing_after_catch_cleanup,
    r#"
        var cleaned = 0
        fun main() {
            try {
                try {
                    throw Exception("x")
                } finally {
                    cleaned += 1
                }
            } catch (e: Exception) {
                println(cleaned)
            }
        }
    "#,
    &["1"]
);

kotlin_run_test!(
    test_throwing_custom_recovery_chain,
    r#"
        fun parseValue(s: String): Int {
            return try {
                s.toInt()
            } catch (e: NumberFormatException) {
                -1
            }
        }
        fun main() {
            println(parseValue("9"))
            println(parseValue("bad"))
        }
    "#,
    &["9", "-1"]
);

kotlin_run_test!(
    test_throwing_chain_preserves_error_class,
    r#"
        fun main() {
            try {
                try {
                    throw IllegalArgumentException("bad")
                } catch (e: Exception) {
                    throw RuntimeException(e)
                }
            } catch (e: RuntimeException) {
                println(e.cause != null)
            }
        }
    "#,
    &["true"]
);

kotlin_run_test!(
    test_throwing_multiple_fails,
    r#"
        fun fail(kind: Int) {
            when (kind) {
                1 -> throw IllegalArgumentException("a")
                2 -> throw IllegalStateException("b")
                else -> throw Exception("c")
            }
        }
        fun main() {
            for (i in 1..3) {
                try {
                    fail(i)
                } catch (e: Exception) {
                    println(e::class.java.simpleName)
                }
            }
        }
    "#,
    &["IllegalArgumentException", "IllegalStateException", "Exception"]
);

kotlin_run_test!(
    test_throwing_with_value_result,
    r#"
        fun recover(v: Int): Int {
            return try {
                if (v == 0) throw Exception("x")
                100 / v
            } catch (e: Exception) {
                0
            }
        }
        fun main() {
            println(recover(4))
            println(recover(0))
        }
    "#,
    &["25", "0"]
);

kotlin_run_test!(
    test_throwing_in_loop_expression,
    r#"
        fun main() {
            var out = 0
            var i = 0
            do {
                try {
                    if (i == 2) throw Exception("x")
                    out += i
                } catch (e: Exception) {
                    out += 10
                }
                i += 1
            } while (i < 4)
            println(out)
        }
    "#,
    &["12"]
);

kotlin_run_test!(
    test_throwing_in_nested_call,
    r#"
        class Boom : Exception("boom")
        fun explode() = throw Boom()
        fun main() {
            try {
                explode()
            } catch (e: Boom) {
                println(e.message)
            }
        }
    "#,
    &["boom"]
);

kotlin_run_test!(
    test_throwing_in_while,
    r#"
        fun main() {
            var i = 0
            while (i < 3) {
                try {
                    if (i == 1) throw Exception("x")
                    println(i)
                } catch (e: Exception) {
                    println("err")
                }
                i += 1
            }
        }
    "#,
    &["0", "err", "2"]
);
