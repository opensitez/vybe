// vybe-test: kotlin/kotlin_collection_grouping/test_grouping_by_index_even_odd_count
// origin: languages/kotlin/tests/kotlin/test_kotlin_collection_grouping.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val grouped = (1..7).withIndex().groupBy { it.index % 2 }
            __check((grouped[0]!!.size).toString(), "4")
            __check((grouped[1]!!.size).toString(), "3")
        }
