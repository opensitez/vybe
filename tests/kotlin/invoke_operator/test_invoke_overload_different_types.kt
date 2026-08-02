// vybe-test: kotlin/invoke_operator/test_invoke_overload_different_types
// origin: languages/kotlin/tests/kotlin/test_invoke_operator.rs

class Printer {
            operator fun invoke(v: Int): String = "i$v"
            operator fun invoke(v: String): String = "s$v"
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val p = Printer()
            __check((p(3)).toString(), "i3")
            __check((p("x")).toString(), "sx")
        }
