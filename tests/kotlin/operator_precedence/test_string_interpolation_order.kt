// vybe-test: kotlin/operator_precedence/test_string_interpolation_order
// origin: languages/kotlin/tests/kotlin/test_operator_precedence.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = 1
            val b = 2
            val c = 3
            __check((a + b * c).toString(), "7")
            __check(("${a + b}*${c}").toString(), "3*3")
        }
