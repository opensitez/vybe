// vybe-test: kotlin/kotlin_collection_grouping/test_grouping_by_lexicographic_case
// origin: languages/kotlin/tests/kotlin/test_kotlin_collection_grouping.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf("a", "A", "b", "B")
            val grouped = values.groupBy { it.lowercase().first() }
            __check((grouped['a']!!.size).toString(), "2")
            __check((grouped['b']!!.size).toString(), "2")
        }
