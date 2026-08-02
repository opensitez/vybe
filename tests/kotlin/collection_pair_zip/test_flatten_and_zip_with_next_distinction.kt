// vybe-test: kotlin/collection_pair_zip/test_flatten_and_zip_with_next_distinction
// origin: languages/kotlin/tests/kotlin/test_collection_pair_zip.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val groups = listOf(listOf(1, 2), listOf(3, 4), listOf(5, 6))
            val zipped = groups.zipWithNext { a, b -> a.last() + b.first() }
            __check((zipped.joinToString(",")).toString(), "5,9")
        }
