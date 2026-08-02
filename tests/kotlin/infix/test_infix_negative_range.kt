// vybe-test: kotlin/infix/test_infix_negative_range
// origin: languages/kotlin/tests/kotlin/test_infix.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val inRange = (-3..3)
__check((0 in inRange).toString(), "true")
__check((4 in inRange).toString(), "false") }
