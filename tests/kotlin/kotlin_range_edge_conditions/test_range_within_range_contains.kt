// vybe-test: kotlin/kotlin_range_edge_conditions/test_range_within_range_contains
// origin: languages/kotlin/tests/kotlin/test_kotlin_range_edge_conditions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val outer = 1..20
            val inner = 5..8
            __check((inner.first() in outer).toString(), "true")
            __check((inner.last() in outer).toString(), "true")
            __check((21 in outer).toString(), "false")
        }
