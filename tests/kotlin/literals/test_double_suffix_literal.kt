// vybe-test: kotlin/literals/test_double_suffix_literal
// origin: languages/kotlin/tests/kotlin/test_literals.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val pi = 3.1415
            val alsoPi = 3.1415d
            __check((pi).toString(), "3.1415")
            __check((alsoPi).toString(), "3.1415")
            __check((alsoPi == pi).toString(), "true")
        }
