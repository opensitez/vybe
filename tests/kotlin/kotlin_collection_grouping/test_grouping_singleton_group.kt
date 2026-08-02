// vybe-test: kotlin/kotlin_collection_grouping/test_grouping_singleton_group
// origin: languages/kotlin/tests/kotlin/test_kotlin_collection_grouping.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf("z")
            val grouped = values.groupBy { it }
            __check((grouped["z"]!!.size).toString(), "1")
            __check((grouped.keys.size).toString(), "1")
        }
