// vybe-test: kotlin/operator_precedence/test_elvis_with_addition
// origin: languages/kotlin/tests/kotlin/test_operator_precedence.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: Int? = null
            __check(((value ?: 5) + 7).toString(), "12")
        }
