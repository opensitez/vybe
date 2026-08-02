// vybe-test: kotlin/type_inference/test_type_inference_in_destructuring
// origin: languages/kotlin/tests/kotlin/test_type_inference.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val pair = Pair(1, "a")
            val (num, text) = pair
            __check((num).toString(), "1")
            __check((text).toString(), "a")
        }
