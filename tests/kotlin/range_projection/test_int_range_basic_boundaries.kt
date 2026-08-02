// vybe-test: kotlin/range_projection/test_int_range_basic_boundaries
// origin: languages/kotlin/tests/kotlin/test_range_projection.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val r = 1..5
            __check((r.first).toString(), "1")
            __check((r.last).toString(), "5")
            __check((r.step).toString(), "1")
            __check((r.count()).toString(), "5")
        }
