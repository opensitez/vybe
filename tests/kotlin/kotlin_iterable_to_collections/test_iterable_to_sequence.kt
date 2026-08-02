// vybe-test: kotlin/kotlin_iterable_to_collections/test_iterable_to_sequence
// origin: languages/kotlin/tests/kotlin/test_kotlin_iterable_to_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val seq = listOf(1, 2, 3).asSequence()
            __check((seq.sum().toString()).toString(), "6")
        }
