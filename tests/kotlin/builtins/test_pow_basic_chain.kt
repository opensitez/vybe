// vybe-test: kotlin/builtins/test_pow_basic_chain
// origin: languages/kotlin/tests/kotlin/test_builtins.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val square = pow(3.0, 2.0)
            val cubic = pow(2.0, 3.0)
            __check((square).toString(), "9")
            __check((cubic).toString(), "8")
        }
