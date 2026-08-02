// vybe-test: kotlin/kotlin_map_query_ops/test_map_get_or_default
// origin: languages/kotlin/tests/kotlin/test_kotlin_map_query_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val m = mapOf("a" to 1, "b" to 2)
            __check((m.getOrDefault("a", 99).toString()).toString(), "1")
            __check((m.getOrDefault("x", 99).toString()).toString(), "99")
            __check((m.getOrElse("b") { 0 }.toString()).toString(), "2")
            __check((m.getOrElse("z") { 77 }.toString()).toString(), "77")
        }
