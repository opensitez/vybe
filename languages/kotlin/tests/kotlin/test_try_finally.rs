kotlin_run_test!(
    test_try_without_exception,
    r#"fun main() { try { println("ok") } finally { println("fin") } }"#,
    &["ok", "fin"]
);

kotlin_run_test!(
    test_try_with_exception_caught,
    r#"fun main() {
        try {
            throw IllegalStateException("x")
        } catch (e: Exception) {
            println(e::class.simpleName)
        } finally {
            println("done")
        }
    }"#,
    &["IllegalStateException", "done"]
);

kotlin_run_test!(
    test_try_with_nested_finally,
    r#"fun main() {
        try {
            try { println("inner") }
            finally { println("inner-finally") }
        } finally { println("outer-finally") }
    }"#,
    &["inner", "inner-finally", "outer-finally"]
);

kotlin_run_test!(
    test_finally_without_exception_order,
    r#"fun main() {
        val out = run {
            var x = ""
            try {
                x += "t"
            } finally {
                x += "f"
            }
            x
        }
        println(out)
    }"#,
    &["tf"]
);

kotlin_run_test!(
    test_return_from_try_finally,
    r#"fun probe(): String {
        try {
            return "try"
        } finally {
            return "finally"
        }
    }
    fun main() { println(probe()) }"#,
    &["finally"]
);

kotlin_run_test!(
    test_return_with_value_and_finally_side_effect,
    r#"fun probe(): Int {
        var x = 0
        try {
            return 1
        } finally {
            x = 2
        }
    }
    fun main() { println(probe()) }"#,
    &["1"]
);

kotlin_run_test!(
    test_mutation_in_finally_observable,
    r#"fun main() {
        var x = 0
        try {
            println("start")
        } finally {
            x = 9
            println(x)
        }
        println(x)
    }"#,
    &["start", "9", "9"]
);

kotlin_run_test!(
    test_finally_runs_on_branch,
    r#"fun main() {
        val v = if (true) {
            try { "yes" } finally { println("fin") }
        } else { "no" }
        println(v)
    }"#,
    &["fin", "yes"]
);

kotlin_run_test!(
    test_finally_after_throw,
    r#"fun main() {
        try {
            try {
                throw RuntimeException("boom")
            } finally {
                println("clean")
            }
        } catch (e: Exception) {
            println("caught")
        }
    }"#,
    &["clean", "caught"]
);

kotlin_run_test!(
    test_catch_multiple_then_finally,
    r#"fun main() {
        try {
            throw IllegalArgumentException()
        } catch (e: IllegalStateException) {
            println("state")
        } catch (e: IllegalArgumentException) {
            println("arg")
        } finally {
            println("done")
        }
    }"#,
    &["arg", "done"]
);

kotlin_run_test!(
    test_try_resource_style_manual,
    r#"fun main() {
        var data = ""
        try {
            data = "open"
        } finally {
            data = data + "-closed"
        }
        println(data)
    }"#,
    &["open-closed"]
);

kotlin_run_test!(
    test_catch_then_else,
    r#"fun main() {
        try {
            println("ok")
        } catch (e: Exception) {
            println("bad")
        } finally {
            println("fin")
        }
    }"#,
    &["ok", "fin"]
);

kotlin_run_test!(
    test_finally_with_loop,
    r#"fun main() {
        var x = 0
        for (i in 0..1) {
            try {
                x += i
            } finally {
                x += 10
            }
        }
        println(x)
    }"#,
    &["22"]
);

kotlin_run_test!(
    test_try_finally_in_function,
    r#"fun run(): Int {
        var x = 0
        try { x = 1 } finally { x = 2 }
        return x
    }
    fun main() { println(run()) }"#,
    &["2"]
);

kotlin_run_test!(
    test_try_finally_with_return_in_catch,
    r#"fun run(): String {
        try {
            throw RuntimeException("x")
        } catch (e: Exception) {
            return "err"
        } finally {
            println("fin")
        }
    }
    fun main() { println(run()) }"#,
    &["fin", "err"]
);

kotlin_run_test!(
    test_nested_try_finally_with_catch,
    r#"fun main() {
        try {
            try {
                throw RuntimeException("x")
            } catch (e: RuntimeException) {
                println("inner")
            } finally {
                println("inner-finally")
            }
        } catch (e: Exception) {
            println("outer")
        } finally {
            println("outer-finally")
        }
    }"#,
    &["inner", "inner-finally", "outer-finally"]
);

kotlin_run_test!(
    test_try_finally_with_break_in_loop,
    r#"fun main() {
        for (i in 1..3) {
            try {
                if (i == 2) break
            } finally {
                println(i)
            }
        }
    }"#,
    &["1", "2"]
);

kotlin_run_test!(
    test_try_finally_with_continue,
    r#"fun main() {
        var x = 0
        for (i in 1..3) {
            try {
                if (i == 2) continue
                x += 1
            } finally {
                x += 10
            }
        }
        println(x)
    }"#,
    &["22"]
);

kotlin_run_test!(
    test_finally_runs_after_return_value_calculation,
    r#"fun f(): Int {
        try {
            return 1 + 1
        } finally {
            println("finally")
        }
    }
    fun main() { println(f()) }"#,
    &["finally", "2"]
);

kotlin_run_test!(
    test_try_finally_without_return,
    r#"fun main() {
        var out = "start"
        try {
            out = "try"
        } finally {
            out += "-fin"
        }
        println(out)
    }"#,
    &["try-fin"]
);

kotlin_run_test!(
    test_finally_and_exception_type,
    r#"fun main() {
        try {
            val x = 1 / 0
            println(x)
        } catch (e: ArithmeticException) {
            println("arith")
        } finally {
            println("done")
        }
    }"#,
    &["arith", "done"]
);

kotlin_run_test!(
    test_multiple_finally_levels,
    r#"fun main() {
        try {
            println("outer-try")
            try {
                println("inner-try")
            } finally {
                println("inner-f")
            }
        } finally {
            println("outer-f")
        }
    }"#,
    &["outer-try", "inner-try", "inner-f", "outer-f"]
);

kotlin_run_test!(
    test_finally_with_unit_return,
    r#"fun run(): Unit {
        try {
            println("u")
        } finally {
            println("f")
        }
    }
    fun main() { run() }"#,
    &["u", "f"]
);

kotlin_run_test!(
    test_finally_with_for_each,
    r#"fun main() {
        val data = intArrayOf(1,2)
        var sum = 0
        data.forEach { v ->
            try {
                sum += v
            } finally {
                sum += 1
            }
        }
        println(sum)
    }"#,
    &["5"]
);

kotlin_run_test!(
    test_finally_masked_exception,
    r#"fun probe(): Int {
        try {
            try {
                throw IllegalStateException("x")
            } finally {
                throw RuntimeException("y")
            }
        } catch (e: RuntimeException) {
            println(e.message)
            return 0
        } finally {
            println("after")
        }
    }
    fun main() { println(probe()) }"#,
    &["y", "after", "0"]
);

kotlin_run_test!(
    test_try_finally_with_local_return,
    r#"fun run(): String {
        var v = ""
        val result = run {
            try {
                "try"
            } finally {
                v = "finally"
            }
        }
        return v + ":" + result
    }
    fun main() { println(run()) }"#,
    &["finally:try"]
);

kotlin_run_test!(
    test_finally_after_no_throw_return,
    r#"fun run(): Int {
        return try {
            7
        } finally {
            println("done")
        }
    }
    fun main() { println(run()) }"#,
    &["done", "7"]
);

kotlin_run_test!(
    test_try_finally_after_if_else,
    r#"fun run(v: Int): Int {
        return if (v > 0) {
            try { v } finally { println("pos") }
        } else {
            try { -v } finally { println("neg") }
        }
    }
    fun main() { println(run(1)); println(run(-2)) }"#,
    &["pos", "1", "neg", "2"]
);

kotlin_run_test!(
    test_finally_with_labelled_block,
    r#"fun main() {
        try {
            println("start")
        } finally { println("finally") }
    }"#,
    &["start", "finally"]
);

kotlin_run_test!(
    test_try_finally_in_tail_position,
    r#"fun f(v: Int): Int = try { v + 1 } finally { println("fin") }
fun main() { println(f(3)) }"#,
    &["fin", "4"]
);

kotlin_run_test!(
    test_finally_on_custom_error_type,
    r#"class Err : Exception()
fun main() {
    try {
        throw Err()
    } catch (e: Err) {
        println("err")
    } finally {
        println("fin")
    }
}"#,
    &["err", "fin"]
);

kotlin_run_test!(
    test_finally_mutate_string_builder,
    r#"fun main() {
        val sb = StringBuilder()
        try {
            sb.append("a")
        } finally {
            sb.append("b")
        }
        println(sb.toString())
    }"#,
    &["ab"]
);

kotlin_run_test!(
    test_finally_always_runs_even_if_caught,
    r#"fun run() {
        try {
            try {
                throw RuntimeException()
            } finally {
                println("inner")
            }
        } catch (e: Exception) {
            println("outer")
        }
    }
    fun main() { run() }"#,
    &["inner", "outer"]
);

kotlin_run_test!(
    test_finally_in_silent_catch,
    r#"fun run() {
        try {
            throw IllegalArgumentException()
        } catch (e: RuntimeException) {
            // ignore
        } finally {
            println("done")
        }
    }
    fun main() { run() }"#,
    &["done"]
);
