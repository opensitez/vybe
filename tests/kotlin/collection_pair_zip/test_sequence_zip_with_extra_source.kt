// vybe-test: kotlin/collection_pair_zip/test_sequence_zip_with_extra_source
// origin: languages/kotlin/tests/kotlin/test_collection_pair_zip.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val zipped = (1..5).asSequence().zip(listOf("a", "b", "c")) { n, s -> "${'$'}n${'$'}s" }.toList()
            __check((zipped.joinToString(",")).toString(), "1a,2b,3c")
        }
