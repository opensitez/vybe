// vybe-test: kotlin/extension_functions/test_extension_function_for_nullable_receiver
// origin: languages/kotlin/tests/kotlin/test_extension_functions.rs

fun Int?.orZero(): Int = this ?: 0

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: Int? = null
            val second: Int? = 7
            __check((value.orZero()).toString(), "0")
            __check((second.orZero()).toString(), "7")
        }
