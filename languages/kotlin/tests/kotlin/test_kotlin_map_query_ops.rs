kotlin_run_cases! {
    test_map_get_or_default => (r##"
        fun main() {
            val m = mapOf("a" to 1, "b" to 2)
            println(m.getOrDefault("a", 99).toString())
            println(m.getOrDefault("x", 99).toString())
            println(m.getOrElse("b") { 0 }.toString())
            println(m.getOrElse("z") { 77 }.toString())
        }
    "##, vec![String::from("1"), String::from("99"), String::from("2"), String::from("77")]),
    test_map_lookup_keys => (r##"
        fun main() {
            val m = mapOf("x" to true, "y" to false)
            println(m.containsKey("x").toString())
            println(m.containsKey("z").toString())
            println(m.containsValue(false).toString())
            println(m.containsValue(true).toString())
        }
    "##, vec![String::from("true"), String::from("false"), String::from("true"), String::from("true")]),
    test_map_merging_plus => (r##"
        fun main() {
            val a = mapOf("a" to 1, "b" to 2)
            val b = mapOf("b" to 3, "c" to 4)
            val merged = a + b
            println(merged["a"].toString())
            println(merged["b"].toString())
            println(merged["c"].toString())
        }
    "##, vec![String::from("1"), String::from("3"), String::from("4")]),
    test_map_minus => (r##"
        fun main() {
            val a = mapOf("a" to 1, "b" to 2, "c" to 3)
            val b = a - listOf("b")
            println(b.size)
            println(b.containsKey("b").toString())
            println(b["c"].toString())
        }
    "##, vec![String::from("2"), String::from("false"), String::from("3")]),
    test_map_filter_keys_values => (r##"
        fun main() {
            val m = mapOf("a" to 1, "b" to 2, "c" to 3)
            val byKeys = m.filterKeys { it == "a" || it == "c" }
            val byValues = m.filterValues { it > 1 }
            println(byKeys.size)
            println(byValues.size)
            println(byKeys["c"].toString())
            println(byValues["a"]?.toString() ?: "null")
        }
    "##, vec![String::from("2"), String::from("2"), String::from("3"), String::from("null")]),
    test_map_map_keys_values => (r##"
        fun main() {
            val m = mapOf("a" to 1, "b" to 2)
            val mappedKeys = m.mapKeys { it.key.uppercase() }
            val mappedValues = m.mapValues { it.value + 10 }
            println(mappedKeys["A"].toString())
            println(mappedValues["b"].toString())
        }
    "##, vec![String::from("1"), String::from("12")]),
    test_map_to_mutable => (r##"
        fun main() {
            val m = mutableMapOf<String, Int>("a" to 1)
            m["b"] = 2
            m.remove("a")
            println(m.size)
            println(m.getOrElse("a") { 0 })
            println(m["b"].toString())
        }
    "##, vec![String::from("1"), String::from("0"), String::from("2")]),
    test_map_get_with_set => (r##"
        fun main() {
            val m = mapOf("a" to 1, "b" to 2)
            val keys = m.keys
            val values = m.values
            println(keys.contains("a").toString())
            println(values.contains(2).toString())
            println(values.contains(99).toString())
        }
    "##, vec![String::from("true"), String::from("true"), String::from("false")]),
    test_map_entries => (r##"
        fun main() {
            val m = mapOf("a" to 1, "b" to 2)
            var sum = 0
            for ((k, v) in m.entries) {
                if (k == "a") sum += v
            }
            println(sum.toString())
            println(m.entries.size)
        }
    "##, vec![String::from("1"), String::from("2")]),
    test_map_replace_and_put => (r##"
        fun main() {
            val m = mutableMapOf("a" to 1)
            m.put("a", 9)
            m["b"] = 2
            println(m["a"].toString())
            println(m.getOrDefault("b", 0).toString())
        }
    "##, vec![String::from("9"), String::from("2")]),
    test_map_clear_and_empty => (r##"
        fun main() {
            val m = mutableMapOf("a" to 1, "b" to 2)
            println(m.isEmpty().toString())
            m.clear()
            println(m.isEmpty().toString())
            println(m.size.toString())
        }
    "##, vec![String::from("false"), String::from("true"), String::from("0")]),
    test_map_to_pairs => (r##"
        fun main() {
            val m = mapOf("a" to 1, "b" to 2)
            val entries = m.toList()
            println(entries.size)
            println(entries[1].first)
            println(entries[1].second.toString())
        }
    "##, vec![String::from("2"), String::from("b"), String::from("2")]),
    test_map_update_if_present => (r##"
        fun main() {
            val m = mutableMapOf("a" to 1)
            m["a"] = (m["a"] ?: 0) + 1
            println(m["a"].toString())
            println(m["x"] ?: -1)
        }
    "##, vec![String::from("2"), String::from("-1")]),
    test_map_compute_like => (r##"
        fun main() {
            val m = mutableMapOf("a" to 1)
            m["a"] = (m["a"] ?: 0) + 10
            m.putIfAbsent("b", 20)
            println(m["a"].toString())
            println(m["b"].toString())
        }
    "##, vec![String::from("11"), String::from("20")]),
}
