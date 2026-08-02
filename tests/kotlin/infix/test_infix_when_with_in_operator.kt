// vybe-test: kotlin/infix/test_infix_when_with_in_operator
// origin: languages/kotlin/tests/kotlin/test_infix.rs

fun score(v: Int): String { return when (v) { in 90..100 -> "A"
in 80..89 -> "B"
else -> "F" } }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((score(95)).toString(), "A")
__check((score(50)).toString(), "F") }
