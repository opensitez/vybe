// vybe-test: kotlin/ordered_collections/test_sorted_map_orders_by_key
// origin: languages/kotlin/tests/kotlin/test_ordered_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = java.util.TreeMap<String, Int>()
            map["b"] = 2
            map["a"] = 1
            map["c"] = 3
            __check((map.keys.joinToString(",")).toString(), "a,b,c")
            __check((map.values.joinToString(",")).toString(), "1,2,3")
        }
