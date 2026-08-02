// vybe-test: kotlin/infix/test_infix_range_includes_boundaries
// origin: languages/kotlin/tests/kotlin/test_infix.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val bounds = 1..5
__check((1 in bounds).toString(), "true")
__check((5 in bounds).toString(), "true")
__check((6 in bounds).toString(), "false") }
