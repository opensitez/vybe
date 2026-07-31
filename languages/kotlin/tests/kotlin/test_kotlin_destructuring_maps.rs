kotlin_run_test!(
    test_destructure_map_entry_in_for_loop,
    r#"
        fun main() {
            val values = mapOf("x" to 1, "y" to 2)
            var total = 0
            for ((k, v) in values) {
                if (k == "x") {
                    total = v
                }
            }
            println(total)
        }
    "#,
    &["1"]
);

kotlin_run_test!(
    test_destructure_map_entry_in_map_call,
    r#"
        fun main() {
            val values = mapOf("a" to 1, "b" to 2)
            val doubled = values.map { (k, v) -> k + v.toString() }
            println(doubled.joinToString(","))
        }
    "#,
    &["a1,b2"]
);

kotlin_run_test!(
    test_mutable_map_entry_rewrite_with_destructure,
    r#"
        fun main() {
            val map = mutableMapOf("a" to 1)
            for ((k, _) in map) {
                map[k] = 5
            }
            println(map["a"])
        }
    "#,
    &["5"]
);

kotlin_run_test!(
    test_destructure_pair_in_list_transform,
    r#"
        fun main() {
            val entries = listOf(Pair("a", 3), Pair("b", 4))
            val keys = entries.map { (k, v) -> "${'$'}k${'$'}v" }
            println(keys.joinToString("|"))
        }
    "#,
    &["a3|b4"]
);

kotlin_run_test!(
    test_destructure_nested_map_entry,
    r#"
        fun main() {
            val groups = mapOf("x" to mapOf("a" to 1), "y" to mapOf("b" to 2))
            var total = 0
            for ((outer, inner) in groups) {
                for ((innerKey, innerValue) in inner) {
                    total += innerValue
                }
            }
            println(total)
        }
    "#,
    &["3"]
);

kotlin_run_test!(
    test_destructure_component_functions_order,
    r#"
        fun main() {
            val e = mapOf("z" to 9).entries.first()
            println(e.component1())
            println(e.component2())
        }
    "#,
    &["z", "9"]
);

kotlin_run_test!(
    test_entry_to_map_transform,
    r#"
        fun main() {
            val map = mapOf("a" to 1, "bb" to 2)
            val out = map
                .toList()
                .associate { (k, v) -> Pair(k + v, v + 1) }
            println(out["a1"])
            println(out["bb2"])
        }
    "#,
    &["2", "3"]
);

kotlin_run_test!(
    test_destructure_when_on_map_entries,
    r#"
        fun main() {
            val values = mapOf("a" to 3, "b" to 0)
            val out = values.map { (k, v) ->
                when (v) {
                    0 -> k + "zero"
                    else -> k + "nz"
                }
            }
            println(out.joinToString(","))
        }
    "#,
    &["azero,bnz"]
);

kotlin_run_test!(
    test_map_entry_destructure_with_default_for_missing,
    r#"
        fun main() {
            val values = mapOf("a" to 1)
            val (a, b) = values.entries.first()
            println(a)
            println(values[a] ?: b)
        }
    "#,
    &["a", "1"]
);

kotlin_run_test!(
    test_entry_iteration_order_preserved_by_map_type,
    r#"
        fun main() {
            val values = linkedMapOf("first" to 1, "second" to 2)
            val first = values.entries.first()
            val last = values.entries.last()
            println(first.key)
            println(last.key)
        }
    "#,
    &["first", "second"]
);

kotlin_run_test!(
    test_destructure_map_entry_to_list_and_sum,
    r#"
        fun main() {
            val values = mapOf("x" to 4, "y" to 5)
            val pair = values.entries.map { (k, v) -> Pair(k.length, v) }.toList()
            val sum = pair.sumOf { it.first + it.second }
            println(sum)
        }
    "#,
    &["11"]
);

kotlin_run_test!(
    test_destructuring_entries_in_filter_step,
    r#"
        fun main() {
            val filtered = mapOf("a" to 1, "b" to 2).filter { (k, v) -> k == "b" || v == 1 }
            println(filtered["a"])
            println(filtered["b"])
        }
    "#,
    &["1", "2"]
);
