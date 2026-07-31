kotlin_run_test!(
    test_block_as_expression_in_variable,
    r#"fun main() { val x = if (true) { 1 + 2 } else { 3 }; println(x) }"#,
    &["3"]
);

kotlin_run_test!(
    test_block_with_local_return,
    r#"fun main() {
        val x = run {
            println("inside")
            5
        }
        println(x)
    }"#,
    &["inside", "5"]
);

kotlin_run_test!(
    test_nested_blocks,
    r#"fun main() { val a = { val b = { 2 + 1 }; b() + 1 }; println(a()) }"#,
    &["4"]
);

kotlin_run_test!(
    test_block_scope_shadowing,
    r#"fun main() { val a = 1; val b = run { val a = 2; a + 1 }; println(a + b) }"#,
    &["4"]
);

kotlin_run_test!(
    test_block_with_early_exit,
    r#"fun f(): Int { val x = run { return 7 }; println("after"); return x }
fun main() { println(f()) }"#,
    &["7"]
);

kotlin_run_test!(
    test_with_expression,
    r#"fun main() { val out = with(StringBuilder()) { append("a"); append("b"); toString() }; println(out) }"#,
    &["ab"]
);

kotlin_run_test!(
    test_apply_expression,
    r#"fun main() { val out = StringBuilder().apply { append("x") }.toString(); println(out) }"#,
    &["x"]
);

kotlin_run_test!(
    test_also_expression,
    r#"fun main() { val out = StringBuilder("a").also { it.append("b") }.toString(); println(out) }"#,
    &["ab"]
);

kotlin_run_test!(
    test_let_expression,
    r#"fun main() { val x = "a".let { it + "b" }; println(x) }"#,
    &["ab"]
);

kotlin_run_test!(
    test_run_expression,
    r#"fun main() { val x = run { val a = 1; val b = 2; a + b }; println(x) }"#,
    &["3"]
);

kotlin_run_test!(
    test_takeif_expression,
    r#"fun main() { val x = 3; val y = x.takeIf { it > 2 }; println(y) }"#,
    &["3"]
);

kotlin_run_test!(
    test_takeif_else,
    r#"fun main() { val x = 1; val y = x.takeIf { it > 2 } ?: 0; println(y) }"#,
    &["0"]
);

kotlin_run_test!(
    test_block_for_mutation_then_use,
    r#"fun main() { val x = run { var a = 1; a += 2; a * 2 }; println(x) }"#,
    &["6"]
);

kotlin_run_test!(
    test_while_with_block,
    r#"fun main() { var x = 0; while (run { x < 2 }) { x += 1 }; println(x) }"#,
    &["2"]
);

kotlin_run_test!(
    test_when_block_result,
    r#"fun main() { val out = when (2) { 1 -> 10; 2 -> run { val a = 5; a + 1 }; else -> 0 }; println(out) }"#,
    &["6"]
);

kotlin_run_test!(
    test_for_block_iterator,
    r#"fun main() { val nums = intArrayOf(1,2,3); var sum = 0; for (x in run { nums }) { sum += x }; println(sum) }"#,
    &["6"]
);

kotlin_run_test!(
    test_block_in_return_type,
    r#"fun f(v: Int): Int = if (v == 0) { 0 } else { val x = v * 2; x }
fun main() { println(f(4)) }"#,
    &["8"]
);

kotlin_run_test!(
    test_block_with_multiple_statements,
    r#"fun main() { val x = run { println("a"); println("b"); 3 }; println(x) }"#,
    &["a", "b", "3"]
);

kotlin_run_test!(
    test_block_with_label,
    r#"fun main() {
        val x = run {
            var i = 0
            if (i == 0) {
                i = 2
            }
            i
        }
        println(x)
    }"#,
    &["2"]
);

kotlin_run_test!(
    test_block_assign_to_property,
    r#"class X { var value = 1 }
fun main() { val x = X(); x.value = run { val a = x.value; a + 4 }; println(x.value) }"#,
    &["5"]
);

kotlin_run_test!(
    test_nested_try_block_expression,
    r#"fun main() {
        val out = run {
            try {
                1
            } catch (e: Exception) {
                0
            }
        }
        println(out)
    }"#,
    &["1"]
);

kotlin_run_test!(
    test_block_with_conditional_cast,
    r#"fun main() { val v: Any = 3; val x = if (v is Int) { v + 1 } else { -1 }; println(x) }"#,
    &["4"]
);

kotlin_run_test!(
    test_block_with_return_from_inner_function,
    r#"fun main() {
        val x = run {
            fun y() = 2
            y() + 3
        }
        println(x)
    }"#,
    &["5"]
);

kotlin_run_test!(
    test_block_in_try_finally,
    r#"fun main() {
        val x = try {
            run {
                1 + 2
            }
        } finally {
            println("done")
        }
        println(x)
    }"#,
    &["done", "3"]
);

kotlin_run_test!(
    test_block_expression_in_map,
    r#"fun main() { val x = mapOf(1 to run { 2 + 3 }, 2 to run { 5 + 6}); println(x[1]!! + x[2]!!) }"#,
    &["16"]
);

kotlin_run_test!(
    test_block_in_sequence,
    r#"fun main() { val x = sequenceOf(1, 2, 3).map { it * run { 2 } }.sum(); println(x) }"#,
    &["12"]
);

kotlin_run_test!(
    test_block_as_parameter,
    r#"fun take(v: Int): Int = v
fun main() { val x = take(run { val a = 1; val b = 2; a + b }); println(x) }"#,
    &["3"]
);

kotlin_run_test!(
    test_block_in_lambda,
    r#"fun main() {
        val fn = { x: Int ->
            run {
                val a = x + 1
                a * 2
            }
        }
        println(fn(3))
    }"#,
    &["8"]
);

kotlin_run_test!(
    test_block_chained_scopes,
    r#"fun main() {
        val x = run {
            val a = 1
            run {
                val b = 2
                a + b
            }
        }
        println(x)
    }"#,
    &["3"]
);

kotlin_run_test!(
    test_block_with_unit,
    r#"fun main() {
        val x = run {
            val a = 1
            val b = 2
            a
        }
        println(x)
    }"#,
    &["1"]
);

kotlin_run_test!(
    test_block_with_try_catch,
    r#"fun main() {
        val x = try {
            run {
                throw IllegalArgumentException()
            }
        } catch (e: Exception) {
            11
        }
        println(x)
    }"#,
    &["11"]
);

kotlin_run_test!(
    test_block_in_class_initializer,
    r#"class K {
        val x = run {
            val a = 1
            val b = 2
            a + b
        }
    }
    fun main() { println(K().x) }"#,
    &["3"]
);

kotlin_run_test!(
    test_block_inside_when_subject,
    r#"fun main() {
        val x = when (run { 3 }) {
            run { 1 + 2 } -> "a"
            else -> "b"
        }
        println(x)
    }"#,
    &["a"]
);

kotlin_run_test!(
    test_block_resulting_boolean,
    r#"fun main() { val x = run { val a = 1; val b = 2; a < b }; println(x) }"#,
    &["true"]
);

kotlin_run_test!(
    test_block_resulting_string,
    r#"fun main() { val x = run { val a = "x"; val b = "y"; a + b }; println(x) }"#,
    &["xy"]
);

kotlin_run_test!(
    test_block_resulting_map,
    r#"fun main() { val x = run { val a = mapOf(1 to 2, 3 to 4); a }; println(x[3]) }"#,
    &["4"]
);

kotlin_run_test!(
    test_block_resulting_list,
    r#"fun main() { val x = run { listOf(1, 2, 3).map { it * 2 } }; println(x[1]) }"#,
    &["4"]
);
