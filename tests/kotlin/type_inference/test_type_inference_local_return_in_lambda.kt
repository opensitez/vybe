// vybe-test: kotlin/type_inference/test_type_inference_local_return_in_lambda
// origin: languages/kotlin/tests/kotlin/test_type_inference.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val f = fun(x: Int): Int {
                return x * 2
            }
            __check((f(6)).toString(), "12")
        }
