// vybe-test: kotlin/step_ranges/test_down_to_progression_count
// origin: languages/kotlin/tests/kotlin/test_step_ranges.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = (5 downTo 1).toList()
            __check((values.joinToString(",")).toString(), "5,4,3,2,1")
        }
