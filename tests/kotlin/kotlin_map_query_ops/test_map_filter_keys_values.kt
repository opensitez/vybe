// vybe-test: kotlin/kotlin_map_query_ops/test_map_filter_keys_values
// origin: languages/kotlin/tests/kotlin/test_kotlin_map_query_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val m = mapOf("a" to 1, "b" to 2, "c" to 3)
            val byKeys = m.filterKeys { it == "a" || it == "c" }
            val byValues = m.filterValues { it > 1 }
            __check((byKeys.size).toString(), "2")
            __check((byValues.size).toString(), "2")
            __check((byKeys["c"].toString()).toString(), "3")
            __check((byValues["a"]?.toString() ?: "null").toString(), "null")
        }
