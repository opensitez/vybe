// vybe-test: kotlin/invoke_operator/test_invoke_with_defaulted_parameter
// origin: languages/kotlin/tests/kotlin/test_invoke_operator.rs

class Formatter {
            operator fun invoke(prefix: String = "[") : String = prefix + "end]"
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val f = Formatter()
            __check((f()).toString(), "[end]")
            __check((f("(")).toString(), "(end]")
        }
