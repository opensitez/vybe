// vybe-test: kotlin/type_inference/test_type_inference_for_mutable_collection
// origin: languages/kotlin/tests/kotlin/test_type_inference.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mutableListOf(1, 2, 3)
            values.add(4)
            __check((values).toString(), "[1, 2, 3, 4]")
        }
