// vybe-test: kotlin/collections_set/test_set_remove_all_operation
// origin: languages/kotlin/tests/kotlin/test_collections_set.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mutableSetOf(1, 2, 3, 4)
            __check((values.removeAll(setOf(2, 4, 6))).toString(), "true")
            __check((values.size).toString(), "2")
            __check((values.contains(2)).toString(), "false")
            __check((values.contains(3)).toString(), "true")
        }
