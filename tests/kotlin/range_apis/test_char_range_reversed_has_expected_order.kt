// vybe-test: kotlin/range_apis/test_char_range_reversed_has_expected_order
// origin: languages/kotlin/tests/kotlin/test_range_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val r = 'c' downTo 'a'
            __check((r.toList().joinToString(",")).toString(), "c,b,a")
            __check((r.first).toString(), "c")
            __check((r.last).toString(), "a")
        }
