// vybe-test: kotlin/kotlin_collection_grouping/test_grouping_multiple_spaces
// origin: languages/kotlin/tests/kotlin/test_kotlin_collection_grouping.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf("x y", "a b", "z z", "x y")
            val grouped = values.groupBy { it[0] }
            __check((grouped['x']!!.size).toString(), "2")
            __check((grouped['a']!!.size).toString(), "1")
        }
