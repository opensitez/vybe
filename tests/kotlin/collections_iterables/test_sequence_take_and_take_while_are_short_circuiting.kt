// vybe-test: kotlin/collections_iterables/test_sequence_take_and_take_while_are_short_circuiting
// origin: languages/kotlin/tests/kotlin/test_collections_iterables.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var mapped = 0
            val seq = sequenceOf(1, 2, 3, 4, 5).map {
                mapped += 1
                it
            }
            val taken = seq.take(3).toList().joinToString(",")
            __check((mapped).toString(), "3")
            val bounded = sequenceOf(1, 2, 3, 4, 5)
                .map { it }
                .takeWhile { it < 4 }
                .toList()
                .joinToString(",")
            __check((bounded).toString(), "1,2,3")
            __check((mapped).toString(), "3")
        }
