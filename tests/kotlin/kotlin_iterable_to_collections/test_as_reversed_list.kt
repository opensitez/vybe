// vybe-test: kotlin/kotlin_iterable_to_collections/test_as_reversed_list
// origin: languages/kotlin/tests/kotlin/test_kotlin_iterable_to_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val m = listOf(1, 2, 3).asReversed()
            __check((m.joinToString(",")).toString(), "3,2,1")
        }
