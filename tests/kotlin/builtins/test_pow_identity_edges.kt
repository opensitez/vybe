// vybe-test: kotlin/builtins/test_pow_identity_edges
// origin: languages/kotlin/tests/kotlin/test_builtins.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((pow(9.0, 0.0)).toString(), "1")
            __check((pow(5.0, 1.0)).toString(), "5")
        }
