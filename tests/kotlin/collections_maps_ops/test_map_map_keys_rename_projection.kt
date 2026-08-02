// vybe-test: kotlin/collections_maps_ops/test_map_map_keys_rename_projection
// origin: languages/kotlin/tests/kotlin/test_collections_maps_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val original = mapOf("a" to 1, "b" to 2)
            val renamed = original.mapKeys { it.key + "!" }
            __check((renamed.keys.joinToString(",")).toString(), "a!,b!")
            __check((renamed["a!"] + renamed["b!"]).toString(), "3")
        }
