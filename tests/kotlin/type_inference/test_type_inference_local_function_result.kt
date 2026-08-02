// vybe-test: kotlin/type_inference/test_type_inference_local_function_result
// origin: languages/kotlin/tests/kotlin/test_type_inference.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            fun f() = 42
            __check((f()).toString(), "42")
        }
