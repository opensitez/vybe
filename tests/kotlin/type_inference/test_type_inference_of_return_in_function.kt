// vybe-test: kotlin/type_inference/test_type_inference_of_return_in_function
// origin: languages/kotlin/tests/kotlin/test_type_inference.rs

fun pick(flag: Boolean) = if (flag) 1 else 0
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((pick(true)).toString(), "1")
            __check((pick(false)).toString(), "0")
        }
