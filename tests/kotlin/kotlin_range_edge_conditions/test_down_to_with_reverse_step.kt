// vybe-test: kotlin/kotlin_range_edge_conditions/test_down_to_with_reverse_step
// origin: languages/kotlin/tests/kotlin/test_kotlin_range_edge_conditions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = 10 downTo 1 step 4
            __check((values.toList().joinToString(",")).toString(), "10,6,2")
        }
