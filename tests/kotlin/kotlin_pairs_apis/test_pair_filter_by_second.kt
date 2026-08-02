// vybe-test: kotlin/kotlin_pairs_apis/test_pair_filter_by_second
// origin: languages/kotlin/tests/kotlin/test_kotlin_pairs_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf(1 to 10, 2 to 3, 3 to 8)
            val filtered = values.filter { it.second > 4 }
            __check((filtered.joinToString(",") { it.first.toString() }).toString(), "1,3")
        }
