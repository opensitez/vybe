// vybe-test: kotlin/collections_set/test_set_of_nested_equality
// origin: languages/kotlin/tests/kotlin/test_collections_set.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val groups = setOf(setOf(1, 2), setOf(1, 2), setOf(2, 1))
            __check((groups.size).toString(), "1")
            __check((groups.contains(setOf(2, 1))).toString(), "true")
        }
