// vybe-test: kotlin/collections_maps_ops/test_map_merge_without_overwrite_keeps_existing
// origin: languages/kotlin/tests/kotlin/test_collections_maps_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val base = linkedMapOf("a" to 1, "b" to 2)
            val extras = mapOf("b" to 20, "c" to 3)
            val merged = extras + base
            __check((merged["a"]).toString(), "1")
            __check((merged["b"]).toString(), "2")
            __check((merged["c"]).toString(), "3")
        }
