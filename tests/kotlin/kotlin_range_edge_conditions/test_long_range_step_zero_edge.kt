// vybe-test: kotlin/kotlin_range_edge_conditions/test_long_range_step_zero_edge
// origin: languages/kotlin/tests/kotlin/test_kotlin_range_edge_conditions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val out = (1L..3L step 2).toList()
            __check((out.joinToString(",")).toString(), "1,3")
        }
