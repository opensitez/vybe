// vybe-test: kotlin/collections_set/test_set_remove_and_contains_all_edges
// origin: languages/kotlin/tests/kotlin/test_collections_set.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mutableSetOf(1, 2, 3)
            __check((values.remove(99)).toString(), "false")
            __check((values.remove(2)).toString(), "true")
            __check((values.containsAll(setOf(1, 2, 3))).toString(), "false")
            __check((values.containsAll(setOf(1, 3))).toString(), "true")
        }
