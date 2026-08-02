// vybe-test: kotlin/collections_set/test_sorted_set_view
// origin: languages/kotlin/tests/kotlin/test_collections_set.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = setOf(4, 2, 1, 3)
            val sorted = values.toSortedSet()
            __check((sorted.first()).toString(), "1")
            __check((sorted.last()).toString(), "4")
            __check((sorted.size).toString(), "4")
        }
