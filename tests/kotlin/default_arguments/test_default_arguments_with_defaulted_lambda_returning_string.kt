// vybe-test: kotlin/default_arguments/test_default_arguments_with_defaulted_lambda_returning_string
// origin: languages/kotlin/tests/kotlin/test_default_arguments.rs

fun render(prefix: String, printer: () -> String = { "x" }): String {
            return prefix + printer()
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((render("a")).toString(), "ax")
            __check((render("a", { "b" })).toString(), "ab")
        }
