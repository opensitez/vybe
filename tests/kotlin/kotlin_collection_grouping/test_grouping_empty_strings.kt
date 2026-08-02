// vybe-test: kotlin/kotlin_collection_grouping/test_grouping_empty_strings
// origin: languages/kotlin/tests/kotlin/test_kotlin_collection_grouping.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf("", "", "x")
            val grouped = values.groupBy { it }
            __check((grouped[""]!!.size).toString(), "2")
            __check((grouped["x"]!!.size).toString(), "1")
        }
