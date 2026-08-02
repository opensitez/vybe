// vybe-test: kotlin/type_inference/test_type_inference_in_function_call_chain
// origin: languages/kotlin/tests/kotlin/test_type_inference.rs

fun first(v: Int): Int = v + 1
        fun second(v: Int): Int = v * 2
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val out = first(second(3))
            __check((out).toString(), "7")
        }
