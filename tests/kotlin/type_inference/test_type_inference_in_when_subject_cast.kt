// vybe-test: kotlin/type_inference/test_type_inference_in_when_subject_cast
// origin: languages/kotlin/tests/kotlin/test_type_inference.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: Any = "abc"
            val len = when (val text = value) {
                is String -> text.length
                else -> 0
            }
            __check((len).toString(), "3")
        }
