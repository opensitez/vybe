// vybe-test: kotlin/collections_sequences/test_sequence_from_list_is_lazy_until_terminal
// origin: languages/kotlin/tests/kotlin/test_collections_sequences.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var built = 0
            val source = listOf(1, 2, 3)
            val seq = source.asSequence().onEach { built += 1 }
            __check(("before").toString(), "before")
            __check((seq.count()).toString(), "3")
            __check(("after").toString(), "after")
            __check((built).toString(), "3")
        }
