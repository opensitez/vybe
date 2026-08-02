// vybe-test: kotlin/kotlin_map_query_ops/test_map_to_mutable
// origin: languages/kotlin/tests/kotlin/test_kotlin_map_query_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val m = mutableMapOf<String, Int>("a" to 1)
            m["b"] = 2
            m.remove("a")
            __check((m.size).toString(), "1")
            __check((m.getOrElse("a") { 0 }).toString(), "0")
            __check((m["b"].toString()).toString(), "2")
        }
