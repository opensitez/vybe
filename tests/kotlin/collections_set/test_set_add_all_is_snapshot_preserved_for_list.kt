// vybe-test: kotlin/collections_set/test_set_add_all_is_snapshot_preserved_for_list
// origin: languages/kotlin/tests/kotlin/test_collections_set.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val source = setOf(1, 2, 3)
            val copied = source.toMutableSet()
            copied.addAll(listOf(3, 4, 5))
            __check((source.toString()).toString(), "[1, 2, 3]")
            __check((copied.toString()).toString(), "[1, 2, 3, 4, 5]")
            __check((source.contains(5)).toString(), "false")
            __check((copied.contains(5)).toString(), "true")
        }
