// vybe-test: kotlin/kotlin_map_query_ops/test_map_update_if_present
// origin: languages/kotlin/tests/kotlin/test_kotlin_map_query_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val m = mutableMapOf("a" to 1)
            m["a"] = (m["a"] ?: 0) + 1
            __check((m["a"].toString()).toString(), "2")
            __check((m["x"] ?: -1).toString(), "-1")
        }
