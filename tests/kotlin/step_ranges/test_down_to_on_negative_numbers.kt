// vybe-test: kotlin/step_ranges/test_down_to_on_negative_numbers
// origin: languages/kotlin/tests/kotlin/test_step_ranges.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(((-1 downTo -3).toList().joinToString(",")).toString(), "-1,-2,-3")
            __check(((-3..-1).toList().joinToString(",")).toString(), "-3,-2,-1")
        }
