// vybe-test: kotlin/extension_functions/test_extension_function_with_receiver_parameter
// origin: languages/kotlin/tests/kotlin/test_extension_functions.rs

fun Int.addWithOffset(offset: Int, label: String): String {
            return label + (this + offset).toString()
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((2.addWithOffset(3, "x")).toString(), "x5")
        }
