// vybe-test: kotlin/if_expressions/test_if_as_statement_changes_outer_variable
// origin: languages/kotlin/tests/kotlin/test_if_expressions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var total = 0
            if (true) {
                total += 2
            } else {
                total += 9
            }
            if (false) {
                total += 9
            } else if (total == 2) {
                total += 3
            }
            __check((total).toString(), "5")
        }
