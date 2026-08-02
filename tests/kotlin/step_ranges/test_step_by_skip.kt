// vybe-test: kotlin/step_ranges/test_step_by_skip
// origin: languages/kotlin/tests/kotlin/test_step_ranges.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = (1..10 step 3).toList()
            __check((values.joinToString(",")).toString(), "1,4,7,10")
        }
