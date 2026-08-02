// vybe-test: kotlin/invoke_operator/test_invoke_infix_style_variable
// origin: languages/kotlin/tests/kotlin/test_invoke_operator.rs

class Box {
            operator fun invoke(v: String): String = "[$v]"
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val f = Box()
            __check((f("x")).toString(), "[x]")
        }
