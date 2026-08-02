// vybe-test: kotlin/try_finally/test_return_with_value_and_finally_side_effect
// origin: languages/kotlin/tests/kotlin/test_try_finally.rs

fun probe(): Int {
        var x = 0
        try {
            return 1
        } finally {
            x = 2
        }
    }
    fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((probe()).toString(), "1") }
