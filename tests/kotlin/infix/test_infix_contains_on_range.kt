// vybe-test: kotlin/infix/test_infix_contains_on_range
// origin: languages/kotlin/tests/kotlin/test_infix.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val r = 1..10
__check((4 in r).toString(), "true")
__check((11 in r).toString(), "false") }
