// vybe-test: kotlin/type_inference/test_type_inference_with_data_class_copy
// origin: languages/kotlin/tests/kotlin/test_type_inference.rs

data class Box(val x: Int)
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val b = Box(1)
            val c = b.copy(x = 2)
            __check((c.x).toString(), "2")
        }
