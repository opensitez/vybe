// vybe-test: kotlin/builtins/test_trig_zero_projection
// origin: languages/kotlin/tests/kotlin/test_builtins.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((sin(0.0)).toString(), "0")
            __check((cos(0.0)).toString(), "1")
            __check((tan(0.0)).toString(), "0")
        }
