// vybe-test: kotlin/builtins/test_sqrt_chain_with_mul_and_div
// origin: languages/kotlin/tests/kotlin/test_builtins.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = sqrt(81.0)
            val half = value / 2.0
            __check((value).toString(), "9")
            __check((half).toString(), "4.5")
        }
