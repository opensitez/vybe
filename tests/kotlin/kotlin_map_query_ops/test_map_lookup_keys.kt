// vybe-test: kotlin/kotlin_map_query_ops/test_map_lookup_keys
// origin: languages/kotlin/tests/kotlin/test_kotlin_map_query_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val m = mapOf("x" to true, "y" to false)
            __check((m.containsKey("x").toString()).toString(), "true")
            __check((m.containsKey("z").toString()).toString(), "false")
            __check((m.containsValue(false).toString()).toString(), "true")
            __check((m.containsValue(true).toString()).toString(), "true")
        }
