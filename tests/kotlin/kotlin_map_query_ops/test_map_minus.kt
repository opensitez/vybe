// vybe-test: kotlin/kotlin_map_query_ops/test_map_minus
// origin: languages/kotlin/tests/kotlin/test_kotlin_map_query_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = mapOf("a" to 1, "b" to 2, "c" to 3)
            val b = a - listOf("b")
            __check((b.size).toString(), "2")
            __check((b.containsKey("b").toString()).toString(), "false")
            __check((b["c"].toString()).toString(), "3")
        }
