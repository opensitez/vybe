// vybe-test: kotlin/kotlin_pairs_apis/test_pair_sort_by_second
// origin: languages/kotlin/tests/kotlin/test_kotlin_pairs_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf("cat" to 3, "bee" to 1, "ant" to 2)
            val sorted = values.sortedBy { it.second }
            __check((sorted.joinToString("|") { it.first.toString() }).toString(), "bee|ant|cat")
        }
