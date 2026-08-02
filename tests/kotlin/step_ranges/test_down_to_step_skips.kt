// vybe-test: kotlin/step_ranges/test_down_to_step_skips
// origin: languages/kotlin/tests/kotlin/test_step_ranges.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = (10 downTo 1 step 4).toList()
            __check((values.joinToString(",")).toString(), "10,6,2")
        }
