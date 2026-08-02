// vybe-test: kotlin/type_inference/test_type_inference_of_range_sum
// origin: languages/kotlin/tests/kotlin/test_type_inference.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = 1..5
            __check((values.sum()).toString(), "15")
        }
