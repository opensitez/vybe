// vybe-test: kotlin/inline_functions/test_inline_generic_identity
// origin: languages/kotlin/tests/kotlin/test_inline_functions.rs

inline fun <T> identity(value: T): T = value

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((identity("kotlin")).toString(), "kotlin")
            __check((identity(9)).toString(), "9")
        }
