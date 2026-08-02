// vybe-test: kotlin/collections_sequences/test_sequence_zip_stops_at_shortest_input
// origin: languages/kotlin/tests/kotlin/test_collections_sequences.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val zipped = (1..5).asSequence().zip(listOf("a", "b", "c")) { n, s -> "$n-$s" }
            __check((zipped.toList().joinToString(",")).toString(), "1-a,2-b,3-c")
            __check((zipped.count()).toString(), "3")
        }
