// vybe-test: kotlin/range_apis/test_until_excludes_end
// origin: languages/kotlin/tests/kotlin/test_range_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val r = 1 until 5
            __check((r.first).toString(), "1")
            __check((r.last).toString(), "4")
            __check((5 in r).toString(), "false")
            __check((4 in r).toString(), "true")
            __check((r.isEmpty()).toString(), "false")
        }
