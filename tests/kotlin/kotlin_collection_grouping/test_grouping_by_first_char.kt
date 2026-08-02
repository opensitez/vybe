// vybe-test: kotlin/kotlin_collection_grouping/test_grouping_by_first_char
// origin: languages/kotlin/tests/kotlin/test_kotlin_collection_grouping.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val words = listOf("apple", "apricot", "banana", "blue")
            val grouped = words.groupBy { it.first() }
            __check((grouped['a']!!.joinToString(",")).toString(), "apple,apricot")
            __check((grouped['b']!!.size).toString(), "2")
        }
