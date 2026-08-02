// vybe-test: kotlin/range_apis/test_int_range_empty_when_step_non_overlap
// origin: languages/kotlin/tests/kotlin/test_range_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val r = 1..5 step 1
            val d = 5 downTo 1
            __check(((1..5 step 0).isEmpty()).toString(), "false")
            __check(((1..5).step(-1).isEmpty()).toString(), "true")
            __check((r.toList().size).toString(), "5")
            __check((d.toList().size).toString(), "5")
        }
