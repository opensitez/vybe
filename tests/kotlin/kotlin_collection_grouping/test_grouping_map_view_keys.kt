// vybe-test: kotlin/kotlin_collection_grouping/test_grouping_map_view_keys
// origin: languages/kotlin/tests/kotlin/test_kotlin_collection_grouping.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val words = listOf("x", "yy", "zzz")
            val grouped = words.groupBy { it.length }
            val keys = grouped.keys.sorted()
            __check((keys.joinToString(",")).toString(), "1,2,3")
            __check((grouped[2]!!.first()).toString(), "yy")
        }
