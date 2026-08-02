// vybe-test: kotlin/kotlin_range_edge_conditions/test_range_projection_is_monotonic
// origin: languages/kotlin/tests/kotlin/test_kotlin_range_edge_conditions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val r = (0..9)
            val down = r.reversed()
            __check((down.first()).toString(), "9")
            __check((down.last()).toString(), "0")
        }
