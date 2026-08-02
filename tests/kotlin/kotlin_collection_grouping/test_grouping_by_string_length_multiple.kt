// vybe-test: kotlin/kotlin_collection_grouping/test_grouping_by_string_length_multiple
// origin: languages/kotlin/tests/kotlin/test_kotlin_collection_grouping.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf("ant", "bear", "cat", "deer", "eel")
            val grouped = values.groupBy(String::length)
            val keys = grouped.keys.toList().sorted()
            __check((keys.joinToString(",")).toString(), "3,4")
            __check((grouped[3]!!.size).toString(), "2")
        }
