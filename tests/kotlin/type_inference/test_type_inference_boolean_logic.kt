// vybe-test: kotlin/type_inference/test_type_inference_boolean_logic
// origin: languages/kotlin/tests/kotlin/test_type_inference.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val x = 1
            val y = x > 0
            __check((y).toString(), "true")
        }
