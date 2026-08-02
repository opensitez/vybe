// vybe-test: kotlin/kotlin_collection_grouping/test_grouping_by_chars
// origin: languages/kotlin/tests/kotlin/test_kotlin_collection_grouping.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf("a,b", "c,d", "a:e")
            val grouped = values.groupBy { it[0] }
            __check((grouped['a']!!.size).toString(), "2")
            __check((grouped['c']!!.first()).toString(), "c,d")
        }
