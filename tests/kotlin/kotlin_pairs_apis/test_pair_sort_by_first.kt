// vybe-test: kotlin/kotlin_pairs_apis/test_pair_sort_by_first
// origin: languages/kotlin/tests/kotlin/test_kotlin_pairs_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf(5 to "five", 2 to "two", 4 to "four")
            val sorted = values.sortedBy { it.first }
            __check((sorted.joinToString("|") { it.first.toString() }).toString(), "2|4|5")
        }
