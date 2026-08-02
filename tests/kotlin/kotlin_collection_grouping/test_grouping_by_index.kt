// vybe-test: kotlin/kotlin_collection_grouping/test_grouping_by_index
// origin: languages/kotlin/tests/kotlin/test_kotlin_collection_grouping.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf("a", "b", "c", "aa")
            val grouped = values.withIndex().groupBy { it.index % 2 }
            __check((grouped[0]!!.map { it.value }.joinToString(",")).toString(), "a,c")
            __check((grouped[1]!!.map { it.value }.joinToString(",")).toString(), "b,aa")
        }
