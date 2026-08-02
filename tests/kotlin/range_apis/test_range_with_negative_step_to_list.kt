// vybe-test: kotlin/range_apis/test_range_with_negative_step_to_list
// origin: languages/kotlin/tests/kotlin/test_range_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val r = 6 downTo 1
            __check((r.step).toString(), "-1")
            __check((r.toList().joinToString(";")).toString(), "6;5;4;3;2;1")
        }
