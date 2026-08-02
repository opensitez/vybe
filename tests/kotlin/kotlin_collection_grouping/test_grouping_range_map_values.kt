// vybe-test: kotlin/kotlin_collection_grouping/test_grouping_range_map_values
// origin: languages/kotlin/tests/kotlin/test_kotlin_collection_grouping.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = (0..6).toList()
            val grouped = values.groupBy { it / 2 }
            __check((grouped[0]!!.joinToString(",")).toString(), "0,1")
            __check((grouped[3]!!.joinToString(",")).toString(), "6")
        }
