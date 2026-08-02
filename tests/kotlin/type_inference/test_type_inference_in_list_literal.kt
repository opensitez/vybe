// vybe-test: kotlin/type_inference/test_type_inference_in_list_literal
// origin: languages/kotlin/tests/kotlin/test_type_inference.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf(1, 2, 3)
            __check((values.sum()).toString(), "6")
        }
