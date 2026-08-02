// vybe-test: kotlin/extension_functions/test_extension_function_can_be_inference_targeted_by_generic_constraint
// origin: languages/kotlin/tests/kotlin/test_extension_functions.rs

fun <T : Number> T.isBig(): Boolean = this.toDouble() > 10.0

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(((3).isBig()).toString(), "false")
            __check(((42).isBig()).toString(), "true")
            __check((0.5.isBig()).toString(), "false")
        }
