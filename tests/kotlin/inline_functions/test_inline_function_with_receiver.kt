// vybe-test: kotlin/inline_functions/test_inline_function_with_receiver
// origin: languages/kotlin/tests/kotlin/test_inline_functions.rs

inline fun String.wrap(prefix: String, suffix: String): String = prefix + this + suffix

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(("k".wrap("<", ">")).toString(), "<k>")
        }
