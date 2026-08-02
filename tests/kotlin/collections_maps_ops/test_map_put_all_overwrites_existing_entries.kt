// vybe-test: kotlin/collections_maps_ops/test_map_put_all_overwrites_existing_entries
// origin: languages/kotlin/tests/kotlin/test_collections_maps_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val base = mutableMapOf("a" to 1, "b" to 2)
            base.putAll(mapOf("b" to 4, "c" to 5))
            __check((base["b"]).toString(), "4")
            __check((base["c"]).toString(), "5")
            __check((base.size).toString(), "3")
        }
