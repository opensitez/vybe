// vybe-test: kotlin/range_apis/test_range_is_empty_false
// origin: languages/kotlin/tests/kotlin/test_range_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val r = 1..1
            __check((r.isEmpty()).toString(), "false")
            __check((r.contains(1)).toString(), "true")
            __check((r.last).toString(), "1")
        }
