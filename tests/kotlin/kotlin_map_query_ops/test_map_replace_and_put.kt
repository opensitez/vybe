// vybe-test: kotlin/kotlin_map_query_ops/test_map_replace_and_put
// origin: languages/kotlin/tests/kotlin/test_kotlin_map_query_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val m = mutableMapOf("a" to 1)
            m.put("a", 9)
            m["b"] = 2
            __check((m["a"].toString()).toString(), "9")
            __check((m.getOrDefault("b", 0).toString()).toString(), "2")
        }
