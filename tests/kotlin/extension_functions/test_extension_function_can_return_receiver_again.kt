// vybe-test: kotlin/extension_functions/test_extension_function_can_return_receiver_again
// origin: languages/kotlin/tests/kotlin/test_extension_functions.rs

fun String.trimAndRepeat(times: Int): String {
            val clean = this.trim()
            return clean.repeat(times)
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(("  x ".trimAndRepeat(3)).toString(), "xxx")
            __check(("z".trimAndRepeat(1)).toString(), "z")
        }
