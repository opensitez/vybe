// vybe-test: kotlin/numeric_types/test_numeric_zero_rules
// origin: languages/kotlin/tests/kotlin/test_numeric_types.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((0).toString(), "0")
            __check((-0).toString(), "0")
            __check((0 + 0).toString(), "0")
            __check((0L).toString(), "0")
            __check((0.0).toString(), "0")
            __check((-0.0).toString(), "0")
            __check((0.0 == -0.0).toString(), "true")
        }
