// vybe-test: kotlin/builtins/test_math_log_exp_roundtrip_small_delta
// origin: languages/kotlin/tests/kotlin/test_builtins.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((round(exp(log(10.0)))).toString(), "10")
            __check((round(exp(ln(2.0) * 3.0))).toString(), "8")
        }
