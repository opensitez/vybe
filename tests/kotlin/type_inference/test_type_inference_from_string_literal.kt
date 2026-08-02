// vybe-test: kotlin/type_inference/test_type_inference_from_string_literal
// origin: languages/kotlin/tests/kotlin/test_type_inference.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val text = "abc"
            __check((text.length).toString(), "3")
        }
