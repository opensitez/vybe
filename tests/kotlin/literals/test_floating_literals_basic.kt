// vybe-test: kotlin/literals/test_floating_literals_basic
// origin: languages/kotlin/tests/kotlin/test_literals.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((1.0).toString(), "1")
            __check((0.5).toString(), "0.5")
            __check((-2.25).toString(), "-2.25")
            __check((3.14).toString(), "3.14")
        }
