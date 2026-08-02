// vybe-test: kotlin/type_inference/test_type_inference_for_lambda_parameter
// origin: languages/kotlin/tests/kotlin/test_type_inference.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val f = { x: Int -> x + 1 }
            __check((f(3)).toString(), "4")
        }
