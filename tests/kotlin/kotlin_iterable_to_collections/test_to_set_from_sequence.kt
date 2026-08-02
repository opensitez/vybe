// vybe-test: kotlin/kotlin_iterable_to_collections/test_to_set_from_sequence
// origin: languages/kotlin/tests/kotlin/test_kotlin_iterable_to_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val out = sequenceOf(1, 2, 2, 3).toSet()
            __check((out.size).toString(), "3")
            __check((out.joinToString(",")).toString(), "1,2,3")
        }
