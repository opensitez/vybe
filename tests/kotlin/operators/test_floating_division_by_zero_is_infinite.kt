// vybe-test: kotlin/operators/test_floating_division_by_zero_is_infinite
// origin: languages/kotlin/tests/kotlin/test_operators.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((10.0 / 0.0).toString(), "Infinity")
            __check((-10.0 / 0.0).toString(), "-Infinity")
            __check(((0.0 / 0.0).isNaN()).toString(), "true")
        }
