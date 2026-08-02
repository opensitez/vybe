// vybe-test: kotlin/kotlin_sequences_generate/test_sequence_of_literals_then_aggregate
// origin: languages/kotlin/tests/kotlin/test_kotlin_sequences_generate.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val s = sequenceOf("a", "b", "c")
            __check((s.joinToString("|")).toString(), "a|b|c")
        }
