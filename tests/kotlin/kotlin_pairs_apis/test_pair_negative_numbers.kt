// vybe-test: kotlin/kotlin_pairs_apis/test_pair_negative_numbers
// origin: languages/kotlin/tests/kotlin/test_kotlin_pairs_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf(-1 to -2, 3 to 4)
            val total = values.map { it.first + it.second }.joinToString(",")
            __check((total).toString(), "-3,7")
            __check((values[0].first < values[1].first).toString(), "true")
        }
