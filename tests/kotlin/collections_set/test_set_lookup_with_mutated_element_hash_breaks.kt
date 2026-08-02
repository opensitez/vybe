// vybe-test: kotlin/collections_set/test_set_lookup_with_mutated_element_hash_breaks
// origin: languages/kotlin/tests/kotlin/test_collections_set.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            data class Item(var id: Int)
            val item = Item(1)
            val values = hashSetOf(item)
            __check((values.contains(item)).toString(), "true")
            item.id = 99
            __check((values.contains(item)).toString(), "false")
        }
