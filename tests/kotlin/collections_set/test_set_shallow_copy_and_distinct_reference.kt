// vybe-test: kotlin/collections_set/test_set_shallow_copy_and_distinct_reference
// origin: languages/kotlin/tests/kotlin/test_collections_set.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mutableSetOf(1, 2, 3)
            val copied = values.toMutableSet()
            copied.add(4)
            values.remove(1)
            __check((values.size).toString(), "2")
            __check((copied.size).toString(), "4")
            __check((values.contains(1)).toString(), "false")
            __check((copied.contains(1)).toString(), "true")
        }
