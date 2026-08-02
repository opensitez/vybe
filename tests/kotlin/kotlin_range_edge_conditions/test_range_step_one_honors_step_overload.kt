// vybe-test: kotlin/kotlin_range_edge_conditions/test_range_step_one_honors_step_overload
// origin: languages/kotlin/tests/kotlin/test_kotlin_range_edge_conditions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = (0..10).step(3)
            __check((values.toList().joinToString(",")).toString(), "0,3,6,9")
        }
