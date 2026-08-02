// vybe-test: kotlin/extension_functions/test_extension_function_with_multiple_parameters
// origin: languages/kotlin/tests/kotlin/test_extension_functions.rs

fun Int.add(value: Int, scale: Int): Int = (this + value) * scale

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((2.add(3, 4)).toString(), "20")
        }
