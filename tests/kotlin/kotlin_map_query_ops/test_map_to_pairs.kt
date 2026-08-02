// vybe-test: kotlin/kotlin_map_query_ops/test_map_to_pairs
// origin: languages/kotlin/tests/kotlin/test_kotlin_map_query_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val m = mapOf("a" to 1, "b" to 2)
            val entries = m.toList()
            __check((entries.size).toString(), "2")
            __check((entries[1].first).toString(), "b")
            __check((entries[1].second.toString()).toString(), "2")
        }
