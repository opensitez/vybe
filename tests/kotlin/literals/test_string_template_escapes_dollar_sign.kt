// vybe-test: kotlin/literals/test_string_template_escapes_dollar_sign
// origin: languages/kotlin/tests/kotlin/test_literals.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(("${'$'}4.99").toString(), "\$4.99")
            __check(("${'$'}{a + 2}").toString(), "\${a + 2}")
            val prefix = "\${prefix}"
            __check((prefix).toString(), "\${prefix}")
        }
