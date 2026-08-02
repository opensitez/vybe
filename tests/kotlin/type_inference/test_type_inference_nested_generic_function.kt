// vybe-test: kotlin/type_inference/test_type_inference_nested_generic_function
// origin: languages/kotlin/tests/kotlin/test_type_inference.rs

fun <T> box(v: T): List<T> = listOf(v)
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val items = box(7)
            __check((items[0]).toString(), "7")
            val text = box("a")
            __check((text[0]).toString(), "a")
        }
