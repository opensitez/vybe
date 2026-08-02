// vybe-test: kotlin/in_keyword/test_in_with_down_to_range
// origin: languages/kotlin/tests/kotlin/test_in_keyword.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((3 in 5 downTo 1).toString(), "true")
            __check((6 in 5 downTo 1).toString(), "false")
        }
