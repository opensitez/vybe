// vybe-test: kotlin/type_inference/test_type_inference_in_try_expression
// origin: languages/kotlin/tests/kotlin/test_type_inference.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val v = try {
                val x = 1 / 1
                x
            } catch (e: Exception) {
                0
            }
            __check((v).toString(), "1")
        }
