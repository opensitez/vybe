// vybe-test: kotlin/kotlin_collection_grouping/test_grouping_after_mapping
// origin: languages/kotlin/tests/kotlin/test_kotlin_collection_grouping.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val words = listOf("one", "two", "three", "four")
            val grouped = words.map { it.uppercase() }.groupBy { it.length }
            __check((grouped[3]!!.joinToString(",")).toString(), "ONE,TWO")
            __check((grouped[4]!!.first()).toString(), "FOUR")
        }
