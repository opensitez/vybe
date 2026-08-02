// vybe-test: kotlin/do_while_control/test_do_while_zero_to_one
// origin: languages/kotlin/tests/kotlin/test_do_while_control.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var i = 0
            val out = kotlin.run {
                i = 1
                i
            }
            __check((i).toString(), "1")
            __check((out).toString(), "1")
        }
