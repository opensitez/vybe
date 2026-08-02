// vybe-test: kotlin/type_inference/test_type_inference_with_if_expression
// origin: languages/kotlin/tests/kotlin/test_type_inference.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = if (true) 1 else 2
            __check((value).toString(), "1")
        }
