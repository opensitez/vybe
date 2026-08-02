// vybe-test: kotlin/type_inference/test_type_inference_for_lambda_without_annotations
// origin: languages/kotlin/tests/kotlin/test_type_inference.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val add = { a: Int, b: Int -> a + b }
            __check((add(1, 2)).toString(), "3")
        }
