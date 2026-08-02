// vybe-test: kotlin/step_ranges/test_until_is_exclusive_end
// origin: languages/kotlin/tests/kotlin/test_step_ranges.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = (1 until 4).toList()
            __check((values.joinToString(";")).toString(), "1;2;3")
        }
