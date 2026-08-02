// vybe-test: kotlin/step_ranges/test_range_reversed_after_to_list
// origin: languages/kotlin/tests/kotlin/test_step_ranges.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = (1..5).toList().asReversed()
            __check((values.joinToString(",")).toString(), "5,4,3,2,1")
        }
