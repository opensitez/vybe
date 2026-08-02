// vybe-test: kotlin/kotlin_collection_grouping/test_grouping_zip_projection
// origin: languages/kotlin/tests/kotlin/test_kotlin_collection_grouping.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val grouped = listOf("aa", "bbb", "cccc", "d").groupBy { it.length }
            val keys = grouped.keys.sorted()
            val out = keys.joinToString(":")
            __check((out).toString(), "1:2:3:4")
        }
