// vybe-test: kotlin/kotlin_collection_grouping/test_grouping_with_mixed_types
// origin: languages/kotlin/tests/kotlin/test_kotlin_collection_grouping.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf("1", "two", "3", "four")
            val grouped = values.groupBy { if (it.all { ch -> ch.isDigit() }) "digit" else "word" }
            __check((grouped["digit"]!!.joinToString(",")).toString(), "1,3")
            __check((grouped["word"]!!.size).toString(), "2")
        }
