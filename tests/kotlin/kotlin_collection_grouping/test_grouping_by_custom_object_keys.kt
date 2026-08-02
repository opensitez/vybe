// vybe-test: kotlin/kotlin_collection_grouping/test_grouping_by_custom_object_keys
// origin: languages/kotlin/tests/kotlin/test_kotlin_collection_grouping.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val data = listOf("ab", "cd", "efg", "hi")
            val grouped = data.groupBy { Pair(it.length, it[0]) }
            __check((grouped[Pair(2, 'a')]!![0]).toString(), "ab")
            __check((grouped[Pair(2, 'c')]!![0]).toString(), "cd")
        }
