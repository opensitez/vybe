// vybe-test: kotlin/collections_sequences/test_sequence_zip_with_collection
// origin: languages/kotlin/tests/kotlin/test_collections_sequences.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val seq = (1..4).asSequence().zip(listOf("a", "b", "c", "d")) { n, s -> "$n-$s" }
            __check((seq.toList().joinToString(",")).toString(), "1-a,2-b,3-c,4-d")
        }
