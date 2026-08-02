// vybe-test: kotlin/scoping_functions/test_take_if_predicate_called_before_returning_value
// origin: languages/kotlin/tests/kotlin/test_scoping_functions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var checks = 0
            val value = 7.takeIf {
                checks++
                it == 7
            }
            __check((value).toString(), "7")
            __check((checks).toString(), "1")
        }
