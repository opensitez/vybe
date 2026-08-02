// vybe-test: kotlin/extension_functions/test_extension_function_on_primitive
// origin: languages/kotlin/tests/kotlin/test_extension_functions.rs

fun Int.incremented(): Int = this + 1

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((3.incremented()).toString(), "4")
        }
