// vybe-test: kotlin/collections_iterables/test_sequence_is_lazy_until_terminal_operation
// origin: languages/kotlin/tests/kotlin/test_collections_iterables.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var called = 0
            val source = sequenceOf(1, 2, 3, 4, 5)
            val transformed = source.map {
                called += 1
                it * 2
            }
            __check((called).toString(), "0")
            val first = transformed.first()
            __check((called).toString(), "1")
            __check((first).toString(), "2")
            val rest = transformed.take(2).toList()
            __check((rest.joinToString(",")).toString(), "4,6")
            __check((called).toString(), "3")
        }
