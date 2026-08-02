// vybe-test: kotlin/range_projection/test_closed_range_contains
// origin: languages/kotlin/tests/kotlin/test_range_projection.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val r = 3..7
            __check((3 in r).toString(), "true")
            __check((8 in r).toString(), "false")
            __check((r.contains(5)).toString(), "true")
            __check((r.contains(2)).toString(), "false")
        }
