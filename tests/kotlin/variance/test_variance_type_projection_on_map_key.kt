// vybe-test: kotlin/variance/test_variance_type_projection_on_map_key
// origin: languages/kotlin/tests/kotlin/test_variance.rs

fun keyCount(map: Map<out String, *>) = map.size
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((keyCount(mapOf("a" to 1, "b" to 2))).toString(), "2")
            __check((keyCount(mapOf("z" to "x"))).toString(), "1")
        }
