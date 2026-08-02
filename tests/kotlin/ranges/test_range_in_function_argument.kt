// vybe-test: kotlin/ranges/test_range_in_function_argument
// origin: languages/kotlin/tests/kotlin/test_ranges.rs

fun containsTarget(range: IntRange, target: Int): Boolean {
            return target in range
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((containsTarget(3..8, 5)).toString(), "true")
            __check((containsTarget(3..8, 9)).toString(), "false")
        }
