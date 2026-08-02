// vybe-test: kotlin/step_ranges/test_range_contains_inclusive_bounds
// origin: languages/kotlin/tests/kotlin/test_step_ranges.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((1 in 1..3).toString(), "true")
            __check((3 in 1..3).toString(), "true")
            __check((4 in 1..3).toString(), "false")
        }
