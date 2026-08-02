// vybe-test: kotlin/infix/test_infix_custom_add
// origin: languages/kotlin/tests/kotlin/test_infix.rs

class Adder(val base: Int) { infix fun plusValue(other: Int): Int = base + other }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((Adder(5) plusValue 4).toString(), "9") }
