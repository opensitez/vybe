// vybe-test: kotlin/collections_set/test_sorted_set_with_string_ordering
// origin: languages/kotlin/tests/kotlin/test_collections_set.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val ordered = sortedSetOf(3, 1, 2)
            __check((ordered.first()).toString(), "1")
            __check((ordered.last()).toString(), "3")
            __check((ordered.joinToString(",")).toString(), "1,2,3")
        }
