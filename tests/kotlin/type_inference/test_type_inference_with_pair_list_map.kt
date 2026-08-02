// vybe-test: kotlin/type_inference/test_type_inference_with_pair_list_map
// origin: languages/kotlin/tests/kotlin/test_type_inference.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val pairs = listOf(Pair(1, "a"), Pair(2, "b"))
            val map = pairs.toMap()
            __check((map[2]).toString(), "b")
        }
