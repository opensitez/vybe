// vybe-test: kotlin/kotlin_sequences_generate/test_sequence_zipwithnext_side_effect
// origin: languages/kotlin/tests/kotlin/test_kotlin_sequences_generate.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = sequenceOf(1, 2, 3, 4).zipWithNext()
            __check((values.toList().joinToString("|") { "${'$'}{it.first}-${'$'}{it.second}" }).toString(), "1-2|2-3|3-4")
        }
