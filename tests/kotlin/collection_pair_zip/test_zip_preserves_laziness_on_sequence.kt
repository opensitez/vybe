// vybe-test: kotlin/collection_pair_zip/test_zip_preserves_laziness_on_sequence
// origin: languages/kotlin/tests/kotlin/test_collection_pair_zip.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val counted = sequenceOf(1, 2).map { it + 1 }
            val zipped = counted.zip(sequenceOf(4, 5, 6)) { a, b -> a + b }
            __check((zipped.joinToString(",")).toString(), "5,7")
        }
