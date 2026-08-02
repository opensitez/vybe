// vybe-test: kotlin/kotlin_map_query_ops/test_map_clear_and_empty
// origin: languages/kotlin/tests/kotlin/test_kotlin_map_query_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val m = mutableMapOf("a" to 1, "b" to 2)
            __check((m.isEmpty().toString()).toString(), "false")
            m.clear()
            __check((m.isEmpty().toString()).toString(), "true")
            __check((m.size.toString()).toString(), "0")
        }
