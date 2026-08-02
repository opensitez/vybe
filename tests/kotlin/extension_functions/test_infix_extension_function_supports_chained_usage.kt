// vybe-test: kotlin/extension_functions/test_infix_extension_function_supports_chained_usage
// origin: languages/kotlin/tests/kotlin/test_extension_functions.rs

infix fun String.tagWith(prefix: String): String = prefix + this

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val result = "world" tagWith "hello-" tagWith "!"
            __check((result).toString(), "!hello-world")
        }
