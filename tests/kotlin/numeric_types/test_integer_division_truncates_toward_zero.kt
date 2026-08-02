// vybe-test: kotlin/numeric_types/test_integer_division_truncates_toward_zero
// origin: languages/kotlin/tests/kotlin/test_numeric_types.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((7 / 3).toString(), "2")
            __check((8 / 4).toString(), "2")
            __check((-7 / 3).toString(), "-2")
            __check((7 / -3).toString(), "-2")
        }
