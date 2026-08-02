// vybe-test: kotlin/kotlin_pairs_triples/test_collection_of_pairs_transform
// origin: languages/kotlin/tests/kotlin/test_kotlin_pairs_triples.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val pairs = listOf("a" to 1, "bb" to 2)
            val sums = pairs.map { it.first.length + it.second }
            __check((sums.joinToString(",")).toString(), "2,4")
        }
