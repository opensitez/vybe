// vybe-test: kotlin/step_ranges/test_until_range_exclusive
// origin: languages/kotlin/tests/kotlin/test_step_ranges.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = (1 until 1).toList()
            val other = (3 until 5).toList()
            __check((values.isEmpty()).toString(), "true")
            __check((other.joinToString(",")).toString(), "3,4")
        }
