// vybe-test: kotlin/type_inference/test_type_inference_with_nullable_list
// origin: languages/kotlin/tests/kotlin/test_type_inference.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values: List<Int>? = listOf(1, 2)
            __check((values?.size).toString(), "2")
        }
