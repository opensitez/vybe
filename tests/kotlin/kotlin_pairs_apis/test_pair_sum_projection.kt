// vybe-test: kotlin/kotlin_pairs_apis/test_pair_sum_projection
// origin: languages/kotlin/tests/kotlin/test_kotlin_pairs_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf(1 to 10, 2 to 20, 3 to 30)
            val total = values.map { it.first + it.second }.sum()
            __check((total).toString(), "66")
        }
