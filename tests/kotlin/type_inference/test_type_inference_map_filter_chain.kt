// vybe-test: kotlin/type_inference/test_type_inference_map_filter_chain
// origin: languages/kotlin/tests/kotlin/test_type_inference.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val sum = listOf(1, 2, 3).filter { it > 1 }.sum()
            __check((sum).toString(), "5")
        }
