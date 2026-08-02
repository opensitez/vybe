// vybe-test: kotlin/math_builtins/test_log_exp_roundtrip_for_e_identity
// origin: languages/kotlin/tests/kotlin/test_math_builtins.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val eRounded = kotlin.math.round(kotlin.math.exp(1.0) * 1000.0) / 1000.0
            val recovered = kotlin.math.ln(kotlin.math.exp(1.0))
            __check((recovered).toString(), "1.0")
            __check((eRounded > 2.7).toString(), "true")
        }
