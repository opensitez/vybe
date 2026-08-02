// vybe-test: kotlin/ordered_collections/test_map_iteration_over_entries_order
// origin: languages/kotlin/tests/kotlin/test_ordered_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = linkedMapOf("first" to 1, "second" to 2, "third" to 3)
            __check((map.entries.joinToString(";") { it.key }).toString(), "first,second,third")
            __check((map.entries.joinToString(";") { it.value.toString() }).toString(), "1,2,3")
        }
