// vybe-test: kotlin/collections_sequences/test_generate_sequence_string_builder
// origin: languages/kotlin/tests/kotlin/test_collections_sequences.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val seq = generateSequence("a") { prev -> if (prev.length < 3) prev + prev else null }
            __check((seq.joinToString(",")).toString(), "a,aa,aaaa")
        }
