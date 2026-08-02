// vybe-test: kotlin/preconditions/test_check_contract_with_boolean_expression
// origin: languages/kotlin/tests/kotlin/test_preconditions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val total = 5 + 6
            check(total == 11) { "sum mismatch" }
            __check((total).toString(), "11")
        }
