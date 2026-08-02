// vybe-test: kotlin/in_keyword/test_in_range_with_step
// origin: languages/kotlin/tests/kotlin/test_in_keyword.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((5 in 1..10 step 2).toString(), "true")
            __check((6 in 1..10 step 2).toString(), "false")
            __check((4 in 10 downTo 1 step 2).toString(), "true")
            __check((5 in 10 downTo 1 step 2).toString(), "false")
        }
