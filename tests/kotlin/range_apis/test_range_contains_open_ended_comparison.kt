// vybe-test: kotlin/range_apis/test_range_contains_open_ended_comparison
// origin: languages/kotlin/tests/kotlin/test_range_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val r = 2..2
            __check((r.first).toString(), "2")
            __check((r.last).toString(), "2")
            __check((r.contains(2)).toString(), "true")
            __check((r.contains(3)).toString(), "false")
        }
