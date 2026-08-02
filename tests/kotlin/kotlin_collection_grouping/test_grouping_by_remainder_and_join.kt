// vybe-test: kotlin/kotlin_collection_grouping/test_grouping_by_remainder_and_join
// origin: languages/kotlin/tests/kotlin/test_kotlin_collection_grouping.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val grouped = listOf(5, 6, 7, 8, 9, 10).groupBy { it % 2 }
            val evenFirst = grouped[0]!![0]
            val oddLast = grouped[1]!![2]
            __check((evenFirst + oddLast).toString(), "16")
        }
