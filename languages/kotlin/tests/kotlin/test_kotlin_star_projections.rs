kotlin_run_cases! {
    test_star_projection_first => (r#"
        fun firstValue(items: List<*>): String {
            if (items.isEmpty()) {
                return "none"
            }
            return items[0].toString()
        }

        fun main() {
            val mixed: List<Any?> = listOf("x", 2, true)
            println(firstValue(mixed))
            println(firstValue(listOf()))
        }
    "#, vec!["x", "none"]),
    test_star_projection_count => (r#"
        fun isPresent(values: List<*>): String {
            return if (values.isNotEmpty()) "has" else "empty"
        }

        fun main() {
            println(isPresent(listOf(1)))
            println(isPresent(listOf<Int?>()))
        }
    "#, vec!["has", "empty"]),
    test_star_projection_map_view => (r#"
        fun firstKey(values: Map<*, *>): String {
            if (values.isEmpty()) {
                return "none"
            }
            return values.keys.iterator().next().toString()
        }

        fun main() {
            println(firstKey(mapOf("a" to 1, "b" to 2)))
            println(firstKey(mapOf<String, Int>()))
        }
    "#, vec!["a", "none"]),
}
