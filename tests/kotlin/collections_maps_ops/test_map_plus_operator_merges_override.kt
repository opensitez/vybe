// vybe-test: kotlin/collections_maps_ops/test_map_plus_operator_merges_override
// origin: languages/kotlin/tests/kotlin/test_collections_maps_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = mapOf("a" to 1, "b" to 2)
            val b = mapOf("b" to 9, "c" to 3)
            val merged = a + b
            __check((merged["a"]).toString(), "1")
            __check((merged["b"]).toString(), "9")
            __check((merged["c"]).toString(), "3")
            __check((merged.size).toString(), "3")
        }
