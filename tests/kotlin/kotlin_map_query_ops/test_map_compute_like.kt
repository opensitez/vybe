// vybe-test: kotlin/kotlin_map_query_ops/test_map_compute_like
// origin: languages/kotlin/tests/kotlin/test_kotlin_map_query_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val m = mutableMapOf("a" to 1)
            m["a"] = (m["a"] ?: 0) + 10
            m.putIfAbsent("b", 20)
            __check((m["a"].toString()).toString(), "11")
            __check((m["b"].toString()).toString(), "20")
        }
