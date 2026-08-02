// vybe-test: kotlin/ordered_collections/test_map_keys_to_list_after_mutation
// origin: languages/kotlin/tests/kotlin/test_ordered_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = linkedMapOf("a" to 1, "b" to 2)
            val keys = map.keys.toMutableList()
            map["c"] = 3
            __check((keys.joinToString(",")).toString(), "a,b")
            __check((map.keys.joinToString(",")).toString(), "a,b,c")
        }
