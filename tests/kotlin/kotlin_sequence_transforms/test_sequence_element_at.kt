// vybe-test: kotlin/kotlin_sequence_transforms/test_sequence_element_at
// origin: languages/kotlin/tests/kotlin/test_kotlin_sequence_transforms.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val seq = sequenceOf(5, 6, 7)
            __check((seq.elementAt(1).toString()).toString(), "6")
            __check((seq.elementAtOrElse(10) { -1 }.toString()).toString(), "-1")
        }
