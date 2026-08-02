// vybe-test: kotlin/kotlin_sequence_transforms/test_sequence_join_to_string
// origin: languages/kotlin/tests/kotlin/test_kotlin_sequence_transforms.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val seq = sequenceOf("a", "b", "c")
            __check((seq.joinToString("-")).toString(), "a-b-c")
        }
