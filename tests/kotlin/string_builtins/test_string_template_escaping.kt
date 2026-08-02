// vybe-test: kotlin/string_builtins/test_string_template_escaping
// origin: languages/kotlin/tests/kotlin/test_string_builtins.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val name = "k"
            val count = 2
            __check(("$name$count").toString(), "k2")
            __check(("${'$'}{name.uppercase()}${'$'}count").toString(), "K2")
        }
