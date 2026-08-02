// vybe-test: kotlin/numeric_types/test_unary_plus_and_unary_minus
// origin: languages/kotlin/tests/kotlin/test_numeric_types.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = 7
            val b = -7
            __check((+a).toString(), "7")
            __check((-a).toString(), "-7")
            __check((+b).toString(), "-7")
            __check((-b).toString(), "7")
        }
