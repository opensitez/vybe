// vybe-test: kotlin/infix/test_infix_to_nested
// origin: languages/kotlin/tests/kotlin/test_infix.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val a = (1 to 2)
val b = (3 to 4)
__check((a.first + b.second).toString(), "5") }
