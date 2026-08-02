// vybe-test: kotlin/builtins/test_trig_odd_even_symmetry_properties
// origin: languages/kotlin/tests/kotlin/test_builtins.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val angle = 0.42
            __check((abs(sin(angle) + sin(-angle)) < 1.0e-12).toString(), "true")
            __check((abs(cos(angle) - cos(-angle)) < 1.0e-12).toString(), "true")
        }
