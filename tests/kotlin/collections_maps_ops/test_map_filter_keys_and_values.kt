// vybe-test: kotlin/collections_maps_ops/test_map_filter_keys_and_values
// origin: languages/kotlin/tests/kotlin/test_collections_maps_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val data = mapOf("apple" to 5, "kiwi" to 3, "pear" to 8)
            val shortKeys = data.filterKeys { it.length < 5 }
            val highValues = data.filterValues { it >= 6 }
            __check((shortKeys.keys.joinToString(",")).toString(), "kiwi")
            __check((highValues.keys.joinToString(",")).toString(), "pear")
        }
