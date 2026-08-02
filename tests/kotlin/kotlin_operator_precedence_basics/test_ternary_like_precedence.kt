// vybe-test: kotlin/kotlin_operator_precedence_basics/test_ternary_like_precedence
// origin: languages/kotlin/tests/kotlin/test_kotlin_operator_precedence_basics.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val x = 5
            val y = 2
            __check((x + y * 2 - 1).toString(), "8")
            __check(((x + y) * (2 - 1)).toString(), "7")
        }
