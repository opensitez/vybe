// vybe-test: kotlin/kotlin_sequence_transforms/test_sequence_any_all
// origin: languages/kotlin/tests/kotlin/test_kotlin_sequence_transforms.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val seq = sequenceOf(1, 2, 3, 4)
            __check((seq.any { it == 3 }.toString()).toString(), "true")
            __check((seq.all { it > 0 }.toString()).toString(), "true")
        }
