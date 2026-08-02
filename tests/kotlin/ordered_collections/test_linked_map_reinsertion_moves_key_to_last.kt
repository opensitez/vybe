// vybe-test: kotlin/ordered_collections/test_linked_map_reinsertion_moves_key_to_last
// origin: languages/kotlin/tests/kotlin/test_ordered_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = LinkedHashMap<String, Int>()
            map["a"] = 1
            map["b"] = 2
            map["a"] = 9
            __check((map.keys.joinToString(",")).toString(), "a,b")
            __check((map["a"]).toString(), "9")
        }
