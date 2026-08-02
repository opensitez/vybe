// vybe-test: kotlin/type_inference/test_type_inference_for_char_range
// origin: languages/kotlin/tests/kotlin/test_type_inference.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val letters = 'a'..'c'
            __check((letters.count()).toString(), "3")
        }
