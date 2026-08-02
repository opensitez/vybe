// vybe-test: kotlin/nullability/test_null_assertion_operator
// origin: languages/kotlin/tests/kotlin/test_nullability.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val user: String? = "Vybe"
            val nonNull = user!!
            __check((nonNull).toString(), "Vybe")
        }
