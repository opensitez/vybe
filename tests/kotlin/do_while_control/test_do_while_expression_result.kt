// vybe-test: kotlin/do_while_control/test_do_while_expression_result
// origin: languages/kotlin/tests/kotlin/test_do_while_control.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var x = 0
            val value = kotlin.run {
                x = 1
                x
            }
            __check((x).toString(), "1")
            __check((value).toString(), "1")
        }
