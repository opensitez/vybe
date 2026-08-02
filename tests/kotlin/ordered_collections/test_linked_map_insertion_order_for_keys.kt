// vybe-test: kotlin/ordered_collections/test_linked_map_insertion_order_for_keys
// origin: languages/kotlin/tests/kotlin/test_ordered_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = LinkedHashMap<String, Int>()
            map["b"] = 2
            map["a"] = 1
            map["c"] = 3
            __check((map.keys.joinToString(",")).toString(), "b,a,c")
        }
