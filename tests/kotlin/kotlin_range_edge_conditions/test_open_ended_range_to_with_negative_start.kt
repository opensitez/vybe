// vybe-test: kotlin/kotlin_range_edge_conditions/test_open_ended_range_to_with_negative_start
// origin: languages/kotlin/tests/kotlin/test_kotlin_range_edge_conditions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val r = -2..2
            __check((r.first).toString(), "-2")
            __check((r.last).toString(), "2")
            __check((r.count()).toString(), "5")
        }
