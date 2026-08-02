// vybe-test: kotlin/type_inference/test_type_inference_with_set
// origin: languages/kotlin/tests/kotlin/test_type_inference.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val set = mutableSetOf(1, 2, 2)
            __check((set.size).toString(), "2")
        }
