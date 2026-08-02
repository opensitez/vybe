// vybe-test: kotlin/preconditions/test_check_and_require_with_same_expression
// origin: languages/kotlin/tests/kotlin/test_preconditions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val v = 3
            require(v > 0)
            check(v == 3)
            __check((v).toString(), "3")
        }
