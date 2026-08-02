// vybe-test: kotlin/collections_maps_ops/test_map_map_values_transform
// origin: languages/kotlin/tests/kotlin/test_collections_maps_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val source = mapOf("a" to 1, "b" to 2)
            val doubled = source.mapValues { (_, value) -> value * 3 }
            __check((doubled["a"]).toString(), "3")
            __check((doubled["b"]).toString(), "6")
        }
