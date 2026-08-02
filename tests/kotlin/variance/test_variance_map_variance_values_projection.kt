// vybe-test: kotlin/variance/test_variance_map_variance_values_projection
// origin: languages/kotlin/tests/kotlin/test_variance.rs

fun collectValues(values: Map<*, out Number>): Int {
            return values.values.sumBy { it.toInt() }
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((collectValues(mapOf("a" to 1, "b" to 2))).toString(), "3")
            __check((collectValues(mapOf("x" to 3L))).toString(), "3")
        }
