kotlin_run_test!(
    test_pair_first_second_access,
    r#"
        fun main() {
            val p = Pair(10, "ok")
            println(p.first)
            println(p.second)
        }
    "#,
    &["10", "ok"]
);

kotlin_run_test!(
    test_pair_to_infix_constructor,
    r#"
        fun main() {
            val p = 4 to "four"
            println(p.first + 1)
            println(p.second.length)
        }
    "#,
    &["5", "4"]
);

kotlin_run_test!(
    test_triple_indexed_fields,
    r#"
        fun main() {
            val t = Triple(1, 2, 3)
            println(t.third)
            println(t.toList().joinToString(","))
        }
    "#,
    &["3", "1,2,3"]
);

kotlin_run_test!(
    test_pair_destructuring_components,
    r#"
        fun main() {
            val (left, right) = Pair("a", 9)
            println(left)
            println(right)
        }
    "#,
    &["a", "9"]
);

kotlin_run_test!(
    test_triple_destructuring_in_function_return,
    r#"
        fun make(): Triple<String, Int, Boolean> {
            return Triple("x", 4, true)
        }

        fun main() {
            val (k, n, b) = make()
            println(k)
            println(n)
            println(b)
        }
    "#,
    &["x", "4", "true"]
);

kotlin_run_test!(
    test_nested_pair_unpacking,
    r#"
        fun main() {
            val outer = Pair(Pair(1, 2), Pair(3, 4))
            val (left, right) = outer
            val (a, b) = left
            val (c, d) = right
            println(a + b + c + d)
        }
    "#,
    &["10"]
);

kotlin_run_test!(
    test_map_entry_as_pair_api,
    r#"
        fun main() {
            val items = mapOf("a" to 1, "b" to 2)
            val first = items.entries.first()
            println(first.key)
            println(first.value)
        }
    "#,
    &["a", "1"]
);

kotlin_run_test!(
    test_collection_of_pairs_transform,
    r#"
        fun main() {
            val pairs = listOf("a" to 1, "bb" to 2)
            val sums = pairs.map { it.first.length + it.second }
            println(sums.joinToString(","))
        }
    "#,
    &["2,4"]
);

kotlin_run_test!(
    test_pair_component_destructure_in_map_loop,
    r#"
        fun main() {
            val map = mapOf("x" to 10, "y" to 20)
            var total = 0
            for ((k, v) in map) {
                total += if (k == "x") v else 0
            }
            println(total)
        }
    "#,
    &["10"]
);

kotlin_run_test!(
    test_triple_component_functions,
    r#"
        fun main() {
            val point = Triple(1, 2, 3)
            println(point.component1())
            println(point.component2())
            println(point.component3())
        }
    "#,
    &["1", "2", "3"]
);

kotlin_run_test!(
    test_pair_plus_custom_concat,
    r#"
        fun main() {
            val a = listOf(1 to 2)
            val b = listOf(3 to 4)
            val combined = a + b
            val out = combined.joinToString(",") { "${it.first}=${it.second}" }
            println(out)
        }
    "#,
    &["1=2,3=4"]
);
