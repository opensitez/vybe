// vybe-test: kotlin/collections/test_nested_array_flat_map_to_depth_one
// origin: languages/kotlin/tests/kotlin/test_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val buckets = arrayOf(arrayOf(1, 2), arrayOf(3), arrayOf(4, 5))
            val flattened = buckets.flatMap { it.toList() }.toTypedArray()
            __check((flattened.joinToString(",")).toString(), "1,2,3,4,5")
            __check((flattened.size).toString(), "5")
        }
