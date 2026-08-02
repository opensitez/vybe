// vybe-test: kotlin/kotlin_collection_grouping/test_grouping_with_null_key
// origin: languages/kotlin/tests/kotlin/test_kotlin_collection_grouping.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf("a", "", "b", "")
            val grouped = values.groupBy { if (it.isEmpty()) null else it }
            __check((grouped[null]!!.size).toString(), "2")
            __check((grouped["a"]!!.size).toString(), "1")
        }
