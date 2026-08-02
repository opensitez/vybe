// vybe-test: kotlin/kotlin_range_edge_conditions/test_range_contains_boundaries
// origin: languages/kotlin/tests/kotlin/test_kotlin_range_edge_conditions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val r = 1..4
            __check((1 in r).toString(), "true")
            __check((4 in r).toString(), "true")
            __check((5 in r).toString(), "false")
        }
