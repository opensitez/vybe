// vybe-test: kotlin/range_apis/test_long_range_progression
// origin: languages/kotlin/tests/kotlin/test_range_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val r = 10L downTo 6L
            __check((r.first).toString(), "10")
            __check((r.last).toString(), "6")
            __check(((r.step).toString()).toString(), "-1")
            __check((r.toList().joinToString(",")).toString(), "10,9,8,7,6")
        }
