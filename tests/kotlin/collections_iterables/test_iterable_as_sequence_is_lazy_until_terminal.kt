// vybe-test: kotlin/collections_iterables/test_iterable_as_sequence_is_lazy_until_terminal
// origin: languages/kotlin/tests/kotlin/test_collections_iterables.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val source = listOf(1, 2, 3, 4)
            var mappedCount = 0
            val seq = source.asSequence().map {
                mappedCount += 1
                it * 2
            }
            __check((mappedCount).toString(), "0")
            val firstTwo = seq.take(2).toList().joinToString(",")
            __check((mappedCount).toString(), "2")
            __check((firstTwo).toString(), "2,4")
        }
