// vybe-test: kotlin/kotlin_collection_grouping/test_grouping_with_large_input
// origin: languages/kotlin/tests/kotlin/test_kotlin_collection_grouping.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = (1..12).toList()
            val grouped = values.groupBy { it % 4 }
            val totalFirst = grouped[0]!!.sum()
            val totalOther = grouped[1]!!.size + grouped[2]!!.size + grouped[3]!!.size
            __check((totalFirst).toString(), "24")
            __check((totalOther).toString(), "9")
        }
