// vybe-test: kotlin/kotlin_map_query_ops/test_map_get_with_set
// origin: languages/kotlin/tests/kotlin/test_kotlin_map_query_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val m = mapOf("a" to 1, "b" to 2)
            val keys = m.keys
            val values = m.values
            __check((keys.contains("a").toString()).toString(), "true")
            __check((values.contains(2).toString()).toString(), "true")
            __check((values.contains(99).toString()).toString(), "false")
        }
