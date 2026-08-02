// vybe-test: kotlin/kotlin_sequence_transforms/test_sequence_flat_map
// origin: languages/kotlin/tests/kotlin/test_kotlin_sequence_transforms.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val seq = sequenceOf(1, 2).flatMap { v -> sequenceOf(v, v * 10) }
            __check((seq.toList().joinToString(",")).toString(), "1,10,2,20")
        }
