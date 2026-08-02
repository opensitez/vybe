// vybe-test: kotlin/kotlin_collection_grouping/test_grouping_by_unicode_char
// origin: languages/kotlin/tests/kotlin/test_kotlin_collection_grouping.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val words = listOf("ß", "ss", "☃")
            val grouped = words.groupBy { it[0] }
            __check((grouped['ß']!!.size).toString(), "1")
            __check((grouped['☃']!!.first()).toString(), "☃")
        }
