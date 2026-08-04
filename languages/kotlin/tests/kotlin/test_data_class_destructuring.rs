kotlin_run_test!(
    test_data_class_top_level_destructure,
    r#"
        data class PairVal(val x: Int, val y: Int)

        fun main() {
            val (x, y) = PairVal(3, 7)
            println(x)
            println(y)
        }
    "#,
    &["3", "7"]
);

kotlin_run_test!(
    test_destructure_in_function_return,
    r#"
        data class Point(val x: Int, val y: Int)

        fun origin(): Point = Point(0, 0)

        fun main() {
            val (x, y) = origin()
            println(x)
            println(y)
        }
    "#,
    &["0", "0"]
);

kotlin_run_test!(
    test_destructure_with_mutation_of_tuple_like_list,
    r#"
        data class Entry(val a: Int, val b: Int)

        fun main() {
            val source = listOf(Entry(1, 2), Entry(3, 4))
            var sum = 0
            for ((left, right) in source) {
                sum += left + right
            }
            println(sum)
        }
    "#,
    &["10"]
);

kotlin_run_test!(
    test_destructure_map_entries,
    r#"
        fun main() {
            val values = mapOf("a" to 1, "b" to 2)
            var total = 0
            for ((key, value) in values) {
                if (key == "a") {
                    total += value
                }
            }
            println(total)
        }
    "#,
    &["1"]
);

kotlin_run_test!(
    test_lambda_destructure_parameters,
    r#"
        data class Item(val id: Int, val label: String)

        fun main() {
            val values = listOf(Item(1, "a"), Item(2, "b"))
            val out = values.joinToString("-") { (id, label) -> "$id:$label" }
            println(out)
        }
    "#,
    &["1:a-2:b"]
);

kotlin_run_test!(
    test_destructure_to_existing_vars,
    r#"
        data class Holder(val left: Int, val right: Int)

        fun main() {
            val left
            val right
            var out = ""
            run {
                val source = Holder(9, 10)
                val (x, y) = source
                out = "$x,$y"
            }
            println(out)
        }
    "#,
    &["9,10"]
);

kotlin_run_test!(
    test_destructure_into_varargs_works_as_tuple_style,
    r#"
        data class TripleValue(val a: Int, val b: Int, val c: Int)

        fun main() {
            val (a, b, c) = TripleValue(1, 2, 3)
            println(a + b + c)
        }
    "#,
    &["6"]
);

kotlin_run_test!(
    test_destructure_with_default_values_is_explicit_constructor,
    r#"
        data class Node(val value: Int = 4, val name: String = "x")

        fun main() {
            val one = Node(name = "n")
            val two = Node(3, "m")
            val (a, b) = one
            val (c, d) = two
            println(a)
            println(b)
            println(c)
            println(d)
        }
    "#,
    &["4", "n", "3", "m"]
);

kotlin_run_test!(
    test_destructure_nested_data_structures,
    r#"
        data class Left(val value: Int)
        data class Right(val other: Left, val label: String)

        fun main() {
            val item = Right(Left(7), "ok")
            val (left, label) = item
            val (value) = left
            println(value)
            println(label)
        }
    "#,
    &["7", "ok"]
);

kotlin_run_test!(
    test_destructure_function_parameter_returns_single_value,
    r#"
        data class SumPair(val left: Int, val right: Int)

        fun combine(a: Int, b: Int): SumPair = SumPair(a + b, a * b)

        fun main() {
            val (sum, product) = combine(4, 5)
            println(sum)
            println(product)
        }
    "#,
    &["9", "20"]
);
