// vybe-test: kotlin/kotlin_collection_grouping/test_grouping_by_digit_chars
// origin: languages/kotlin/tests/kotlin/test_kotlin_collection_grouping.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf("a1", "b2", "c3", "d4")
            val grouped = values.groupBy { it.last() }
            __check((grouped['1']!!.first()).toString(), "a1")
            __check((grouped['4']!!.first()).toString(), "d4")
        }
