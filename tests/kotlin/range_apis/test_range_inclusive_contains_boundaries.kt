// vybe-test: kotlin/range_apis/test_range_inclusive_contains_boundaries
// origin: languages/kotlin/tests/kotlin/test_range_apis.rs

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
            __check((1 in r).toString(), "true")
            __check((5 in r).toString(), "true")
            __check((6 in r).toString(), "false")
        }
