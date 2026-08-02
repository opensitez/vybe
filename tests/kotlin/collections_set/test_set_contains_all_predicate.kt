// vybe-test: kotlin/collections_set/test_set_contains_all_predicate
// origin: languages/kotlin/tests/kotlin/test_collections_set.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = setOf(1, 2, 3, 4)
            __check((values.containsAll(listOf(1, 4))).toString(), "true")
            __check((values.containsAll(listOf(1, 6))).toString(), "false")
        }
