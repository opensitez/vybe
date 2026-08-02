// vybe-test: kotlin/collection_pair_zip/test_zip_aligns_minimum_length
// origin: languages/kotlin/tests/kotlin/test_collection_pair_zip.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val left = listOf(1, 2, 3, 4)
            val right = listOf("a", "b")
            __check((left.zip(right).joinToString("|") { "${'$'}{it.first}${'$'}{it.second}" }).toString(), "1a|2b")
            __check((left.zip(right).size).toString(), "2")
        }
