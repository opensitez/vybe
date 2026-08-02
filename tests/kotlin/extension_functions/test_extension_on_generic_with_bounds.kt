// vybe-test: kotlin/extension_functions/test_extension_on_generic_with_bounds
// origin: languages/kotlin/tests/kotlin/test_extension_functions.rs

fun <T : Number> T.asIntText(): Int = this.toInt()

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((4.9.asIntText()).toString(), "4")
            __check((7.asIntText()).toString(), "7")
        }
