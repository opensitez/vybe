// vybe-test: kotlin/kotlin_iterable_to_collections/test_to_sorted_set
// origin: languages/kotlin/tests/kotlin/test_kotlin_iterable_to_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val out = listOf(3, 1, 2).toSortedSet()
            __check((out.joinToString(",")).toString(), "1,2,3")
        }
