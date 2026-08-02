// vybe-test: kotlin/kotlin_pairs_apis/test_pair_duplicate_key_keeps_last
// origin: languages/kotlin/tests/kotlin/test_kotlin_pairs_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val source = listOf("x" to 1, "x" to 9)
            val map = source.toMap()
            __check((map["x"]).toString(), "9")
            __check((map.size).toString(), "1")
        }
