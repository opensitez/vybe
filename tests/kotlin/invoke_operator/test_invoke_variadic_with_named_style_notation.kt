// vybe-test: kotlin/invoke_operator/test_invoke_variadic_with_named_style_notation
// origin: languages/kotlin/tests/kotlin/test_invoke_operator.rs

class Tagger {
            operator fun invoke(prefix: String, value: String = "x"): String = prefix + value
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val t = Tagger()
            __check((t("a")).toString(), "ax")
            __check((t(prefix = "b", value = "y")).toString(), "by")
        }
