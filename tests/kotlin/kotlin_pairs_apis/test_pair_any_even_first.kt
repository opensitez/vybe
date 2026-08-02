// vybe-test: kotlin/kotlin_pairs_apis/test_pair_any_even_first
// origin: languages/kotlin/tests/kotlin/test_kotlin_pairs_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf(1 to 0, 2 to 9, 3 to 8)
            __check((values.any { it.first % 2 == 0 }).toString(), "true")
            __check((values.none { it.second < 0 }).toString(), "true")
        }
