kotlin_run_test!(
    test_comment_after_value_expression,
    r#"
        fun main() {
            val x = 1 + // inline comment
            2
            println(x)
        }
    "#,
    &["3"]
);

kotlin_run_test!(
    test_block_comment_in_arithmetic,
    r#"
        fun main() {
            val x = 1 /* comment */ + 2
            println(x)
        }
    "#,
    &["3"]
);

kotlin_run_test!(
    test_block_comment_with_spacing,
    r#"
        fun main() {
            val y = (1 /*a*/) + (2 /*b*/)
            println(y)
        }
    "#,
    &["3"]
);

kotlin_run_test!(
    test_comment_after_else_keyword,
    r#"
        fun main() {
            if (false) {
                println(0)
            } else { // alternate branch
                println(1)
            }
        }
    "#,
    &["1"]
);

kotlin_run_test!(
    test_comment_line_separates_statements,
    r#"
        fun main() {
            val base = 4
            // bump value
            val value = base + 1
            println(value)
        }
    "#,
    &["5"]
);

kotlin_run_test!(
    test_comment_inside_function_signature,
    r#"
        fun total( // comment in signature
            a: Int,
            b: Int
        ): Int = a + b

        fun main() {
            println(total(2, 3))
        }
    "#,
    &["5"]
);

kotlin_run_test!(
    test_comment_after_property,
    r#"
        class Box {
            val value = 11 // stored value
        }

        fun main() {
            println(Box().value)
        }
    "#,
    &["11"]
);

kotlin_run_test!(
    test_comment_after_semicolon,
    r#"
        fun main() {
            val a = 10;
            val b = 5 // trailing comment
            println(a + b)
        }
    "#,
    &["15"]
);

kotlin_run_test!(
    test_comment_in_string_is_literal,
    r#"
        fun main() {
            val text = "// not a comment"
            val other = "a /* block */ b"
            println(text)
            println(other)
        }
    "#,
    &["// not a comment", "a /* block */ b"]
);

kotlin_run_test!(
    test_comment_before_logic,
    r#"
        fun main() {
            val first = 3
            val second = 4
            // branch comment
            println(first + second)
        }
    "#,
    &["7"]
);

kotlin_run_test!(
    test_doc_comment_function,
    r#"
        /** doc comment */
        fun add(a: Int, b: Int) = a + b

        fun main() {
            println(add(1, 2))
        }
    "#,
    &["3"]
);

kotlin_run_test!(
    test_comment_between_type_and_initializer,
    r#"
        fun main() {
            val base: Int // type comment
            = 6
            println(base)
        }
    "#,
    &["6"]
);

kotlin_run_test!(
    test_comment_inside_block_run,
    r#"
        fun main() {
            val result = run {
                /*prepare*/
                4
            }
            println(result)
        }
    "#,
    &["4"]
);

kotlin_run_test!(
    test_multiple_block_comments,
    r#"
        fun main() {
            val value = 2 + /*a*/ 3 + /*b*/ 4
            println(value)
        }
    "#,
    &["9"]
);

kotlin_run_test!(
    test_semicolon_and_comment_chain,
    r#"
        fun main() {
            val a = 1; val b = 2; // pair
            println(a + b)
        }
    "#,
    &["3"]
);

kotlin_run_test!(
    test_comment_between_when_branches,
    r#"
        fun main() {
            val out = when (1) {
                1 -> 10 // first
                2 -> 20
                else -> 30
            }
            println(out)
        }
    "#,
    &["10"]
);

kotlin_run_test!(
    test_comment_after_binary_line,
    r#"
        fun main() {
            val total = 10 + // plus
                20 + // plus
                30
            println(total)
        }
    "#,
    &["60"]
);

kotlin_run_test!(
    test_comment_between_class_members,
    r#"
        class Pairish {
            val a = 1
            // separator
            val b = 2
        }

        fun main() {
            val p = Pairish()
            println(p.a + p.b)
        }
    "#,
    &["3"]
);

kotlin_run_test!(
    test_comment_after_if_condition,
    r#"
        fun main() {
            if (true) // condition comment
            {
                println(1)
            }
        }
    "#,
    &["1"]
);

kotlin_run_test!(
    test_comment_in_while_loop,
    r#"
        fun main() {
            var sum = 0
            var i = 0
            while (i < 2) {
                sum += i
                // keep loop
                i += 1
            }
            println(sum)
        }
    "#,
    &["1"]
);

kotlin_run_test!(
    test_comment_in_for_loop,
    r#"
        fun main() {
            var c = 0
            for (i in 1..2) {
                // count
                c += i
            }
            println(c)
        }
    "#,
    &["3"]
);

kotlin_run_test!(
    test_comment_in_lambda_body,
    r#"
        fun main() {
            val f = { x: Int ->
                // lambda body
                x + 1
            }
            println(f(2))
        }
    "#,
    &["3"]
);

kotlin_run_test!(
    test_comment_inside_raw_string_body,
    r#"
        fun main() {
            val text = """line1
// not parsed as comment
line3"""
            println(text.lines().size)
        }
    "#,
    &["3"]
);

kotlin_run_test!(
    test_comment_next_to_operators,
    r#"
        fun main() {
            val a = 8/*c*/+4
            val b = 2/*c*/+3
            println(a - b)
        }
    "#,
    &["7"]
);

kotlin_run_test!(
    test_comment_between_properties,
    r#"
        class Holder {
            val first = 1
            // comment line
            val second = 2
            val third = 3
        }

        fun main() {
            val h = Holder()
            println(h.first + h.second + h.third)
        }
    "#,
    &["6"]
);

kotlin_run_test!(
    test_comment_between_annotations,
    r#"
        fun main() {
            val count: Int // typed
            = 9
            println(count)
        }
    "#,
    &["9"]
);

kotlin_run_test!(
    test_comment_on_object_member,
    r#"
        fun main() {
            val obj = object {
                val one = 1 // member comment
            }
            println(obj.one)
        }
    "#,
    &["1"]
);

kotlin_run_test!(
    test_comment_before_value,
    r#"
        fun main() {
            val value = 12
            // compute result
            println(value)
        }
    "#,
    &["12"]
);

kotlin_run_test!(
    test_comment_before_else_branch,
    r#"
        fun main() {
            val out = if (false)
                0
            else
                2 // else side
            println(out)
        }
    "#,
    &["2"]
);

kotlin_run_test!(
    test_comment_in_run_block,
    r#"
        fun main() {
            val x = run {
                // compute in block
                5
            }
            println(x)
        }
    "#,
    &["5"]
);

kotlin_run_test!(
    test_comment_after_top_level_decl,
    r#"
        fun main() {
            val prefix = 2
            val suffix = 3
            println(prefix * suffix)
        }
    "#,
    &["6"]
);

kotlin_run_test!(
    test_comment_in_trailing_expression,
    r#"
        fun main() {
            val value = (1 + 2) // addition
            println(value)
        }
    "#,
    &["3"]
);

kotlin_run_test!(
    test_comment_after_for_token,
    r#"
        fun main() {
            val nums = listOf(1, 2, 3)
            var sum = 0
            for (n in nums) { // iterate
                sum += n
            }
            println(sum)
        }
    "#,
    &["6"]
);
