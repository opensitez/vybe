// vybe-test: kotlin/infix/test_infix_double_to
// origin: languages/kotlin/tests/kotlin/test_infix.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val pair = 2 to 4
__check((pair.first + pair.second).toString(), "6") }
