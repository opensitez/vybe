// vybe-test: kotlin/range_apis/test_int_range_reduce_via_terminal
// origin: languages/kotlin/tests/kotlin/test_range_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val r = 2..5
            val value = r.reduce { acc, value -> acc * value }
            __check((value).toString(), "120")
        }
