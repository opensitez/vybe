// vybe-test: kotlin/range_apis/test_range_singleton_down_to
// origin: languages/kotlin/tests/kotlin/test_range_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val r = 1 downTo 1
            __check((r.toList().joinToString(",")).toString(), "1")
            __check((r.isEmpty()).toString(), "false")
            __check((r.first).toString(), "1")
            __check((r.last).toString(), "1")
        }
