// vybe-test: kotlin/type_inference/test_type_inference_from_constructor_call
// origin: languages/kotlin/tests/kotlin/test_type_inference.rs

class Box(val value: Int)
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val b = Box(7)
            __check((b.value).toString(), "7")
        }
