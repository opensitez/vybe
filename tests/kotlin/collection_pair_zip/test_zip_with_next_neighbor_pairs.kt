// vybe-test: kotlin/collection_pair_zip/test_zip_with_next_neighbor_pairs
// origin: languages/kotlin/tests/kotlin/test_collection_pair_zip.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = listOf(1, 2, 3, 4)
            val pairs = nums.zipWithNext().joinToString("|") { "${'$'}{it.first}:${'$'}{it.second}" }
            __check((pairs).toString(), "1:2|2:3|3:4")
        }
