// vybe-test: kotlin/type_inference/test_type_inference_with_elvis
// origin: languages/kotlin/tests/kotlin/test_type_inference.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a: Int? = null
            val b: Int = a ?: 5
            __check((b).toString(), "5")
        }
