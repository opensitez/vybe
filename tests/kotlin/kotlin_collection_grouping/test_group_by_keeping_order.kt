// vybe-test: kotlin/kotlin_collection_grouping/test_group_by_keeping_order
// origin: languages/kotlin/tests/kotlin/test_kotlin_collection_grouping.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf("bb", "a", "c", "aa", "b")
            val grouped = values.groupBy { it.length }
            __check((grouped[1]!!.joinToString("|")).toString(), "a|c|b")
            __check((grouped[2]!!.joinToString("|")).toString(), "bb|aa")
        }
