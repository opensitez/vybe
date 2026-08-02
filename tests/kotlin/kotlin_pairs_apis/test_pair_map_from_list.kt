// vybe-test: kotlin/kotlin_pairs_apis/test_pair_map_from_list
// origin: languages/kotlin/tests/kotlin/test_kotlin_pairs_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val source = listOf("a" to 1, "b" to 2)
            val map = source.toMap()
            __check((map["a"]).toString(), "1")
            __check((map["b"]).toString(), "2")
        }
