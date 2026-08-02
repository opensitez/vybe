// vybe-test: kotlin/infix/test_infix_in_string_search
// origin: languages/kotlin/tests/kotlin/test_infix.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val text = "kotlin"
__check(("li" in text).toString(), "true")
__check(("zz" in text).toString(), "false") }
