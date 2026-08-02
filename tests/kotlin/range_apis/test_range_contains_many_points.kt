// vybe-test: kotlin/range_apis/test_range_contains_many_points
// origin: languages/kotlin/tests/kotlin/test_range_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val r = 10 downTo 1
            __check((10 in r).toString(), "true")
            __check((1 in r).toString(), "true")
            __check((0 in r).toString(), "false")
            __check((11 in r).toString(), "false")
        }
