// vybe-test: kotlin/type_inference/test_type_inference_for_generic_function_output
// origin: languages/kotlin/tests/kotlin/test_type_inference.rs

fun <T> id(v: T): T = v
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val x = id(4)
            val y = id("z")
            __check((x).toString(), "4")
            __check((y).toString(), "z")
        }
