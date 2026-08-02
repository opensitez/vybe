// vybe-test: kotlin/builtins/test_finite_and_infinite_detection_for_division_quirks
// origin: languages/kotlin/tests/kotlin/test_builtins.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(((1.0 / 3.0).isFinite()).toString(), "true")
            __check(((1.0 / 0.0).isInfinite()).toString(), "true")
            __check(((-1.0 / 0.0).isInfinite()).toString(), "true")
            __check(((0.0 / 0.0).isNaN()).toString(), "true")
        }
