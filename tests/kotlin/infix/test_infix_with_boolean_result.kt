// vybe-test: kotlin/infix/test_infix_with_boolean_result
// origin: languages/kotlin/tests/kotlin/test_infix.rs

class Window { infix fun contains(value: Int): Boolean { return value % 2 == 0 } }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val w = Window()
__check((w contains 8).toString(), "true")
__check((w contains 9).toString(), "false") }
