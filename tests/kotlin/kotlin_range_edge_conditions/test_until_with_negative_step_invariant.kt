// vybe-test: kotlin/kotlin_range_edge_conditions/test_until_with_negative_step_invariant
// origin: languages/kotlin/tests/kotlin/test_kotlin_range_edge_conditions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = 5 until 0
            __check((values.count()).toString(), "0")
            __check((values.toList().isEmpty()).toString(), "true")
        }
