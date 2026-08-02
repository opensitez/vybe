// vybe-test: kotlin/literals/test_string_basic_literals
// origin: languages/kotlin/tests/kotlin/test_literals.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(("plain").toString(), "plain")
            __check(("").toString(), "")
            __check(("x").toString(), "x")
        }
