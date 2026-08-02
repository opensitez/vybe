// vybe-test: kotlin/ranges/test_coerce_in_range_boundaries
// origin: languages/kotlin/tests/kotlin/test_ranges.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val allowed = 3..8
            __check((1.coerceIn(allowed)).toString(), "3")
            __check((5.coerceIn(allowed)).toString(), "5")
            __check((9.coerceIn(allowed)).toString(), "8")
        }
