// vybe-test: kotlin/numeric_types/test_remainder_with_floating_values
// origin: languages/kotlin/tests/kotlin/test_numeric_types.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((7.5 % 2.0).toString(), "1.5")
            __check((8.2 % 2.0).toString(), "0.2")
            __check((-7.5 % 2.0).toString(), "-1.5")
        }
