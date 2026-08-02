// vybe-test: kotlin/ranges/test_range_contains_uses_open_ended_bounds_expressions
// origin: languages/kotlin/tests/kotlin/test_ranges.rs

fun clamp(value: Int, base: Int, width: Int): Boolean {
            return value in (base until (base + width))
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((clamp(3, 0, 3)).toString(), "true")
            __check((clamp(2, -2, 4)).toString(), "true")
            __check((clamp(3, -2, 4)).toString(), "false")
        }
