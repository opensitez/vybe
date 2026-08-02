// vybe-test: kotlin/kotlin_collection_grouping/test_grouping_by_empty_string_values
// origin: languages/kotlin/tests/kotlin/test_kotlin_collection_grouping.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val grouped = listOf("", "a", "", "bb", "", "bbb").groupBy { it.length }
            __check((grouped[0]!!.size).toString(), "3")
            __check((grouped[3]!!.size).toString(), "1")
        }
