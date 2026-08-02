// vybe-test: kotlin/range_apis/test_range_last_of_down_to_step
// origin: languages/kotlin/tests/kotlin/test_range_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val r = 9 downTo 3 step 3
            __check((r.toList().joinToString(",")).toString(), "9,6,3")
        }
