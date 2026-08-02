// vybe-test: kotlin/kotlin_pairs_apis/test_triple_in_map
// origin: languages/kotlin/tests/kotlin/test_kotlin_pairs_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mapOf("row" to Triple(1, 2, 3), "other" to Triple(4, 5, 6))
            __check((values["row"]?.first).toString(), "1")
            __check((values["other"]?.third).toString(), "6")
        }
