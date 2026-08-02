// vybe-test: kotlin/ranges/test_range_step_property_is_intprogression_value
// origin: languages/kotlin/tests/kotlin/test_ranges.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val range = 1..10 step 3
            __check((range.step).toString(), "3")
            __check((range.first).toString(), "1")
            __check((range.last).toString(), "10")
        }
