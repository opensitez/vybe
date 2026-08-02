// vybe-test: kotlin/throwing_recovery/test_throwing_in_expression_assignment
// origin: languages/kotlin/tests/kotlin/test_throwing_recovery.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val x = try {
                throw Exception("no")
            } catch (e: Exception) {
                11
            }
            __check((x).toString(), "11")
        }
