// vybe-test: kotlin/collections_maps_ops/test_map_or_empty_for_nullable_source
// origin: languages/kotlin/tests/kotlin/test_collections_maps_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val maybe: Map<String, Int>? = null
            val safe = maybe.orEmpty()
            __check((safe.isEmpty()).toString(), "true")
            __check((safe.size).toString(), "0")
        }
