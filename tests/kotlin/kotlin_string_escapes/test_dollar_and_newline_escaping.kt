// vybe-test: kotlin/kotlin_string_escapes/test_dollar_and_newline_escaping
// origin: languages/kotlin/tests/kotlin/test_kotlin_string_escapes.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(("dollar:" + "\$").toString(), "dollar:\$")
            __check(("slash-n:" + "\\n").toString(), "slash-n:\\n")
        }
