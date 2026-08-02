// vybe-test: kotlin/variance/test_variance_map_kotlin_read_projection
// origin: languages/kotlin/tests/kotlin/test_variance.rs

fun readAny(map: Map<String, out Number>): String = map.values.joinToString("-") { it.toString() }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((readAny(mapOf("a" to 1))).toString(), "1")
            __check((readAny(mapOf("b" to 2.0))).toString(), "2.0")
        }
