// vybe-test: kotlin/infix/test_infix_custom_text
// origin: languages/kotlin/tests/kotlin/test_infix.rs

class Verb { infix fun shout(other: String): String = other + other }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((Verb() shout "go").toString(), "gogo") }
