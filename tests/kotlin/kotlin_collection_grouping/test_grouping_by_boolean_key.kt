// vybe-test: kotlin/kotlin_collection_grouping/test_grouping_by_boolean_key
// origin: languages/kotlin/tests/kotlin/test_kotlin_collection_grouping.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf(1, 2, 3, 4, 5)
            val grouped = values.groupBy { it > 3 }
            __check((grouped[true]!!.joinToString(",")).toString(), "4,5")
            __check((grouped[false]!!.joinToString(",")).toString(), "1,2,3")
        }
