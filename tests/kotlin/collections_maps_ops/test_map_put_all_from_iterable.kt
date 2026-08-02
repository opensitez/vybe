// vybe-test: kotlin/collections_maps_ops/test_map_put_all_from_iterable
// origin: languages/kotlin/tests/kotlin/test_collections_maps_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val base = mutableMapOf("a" to 1)
            base.putAll(listOf("b" to 2, "c" to 3))
            __check((base["a"] + base["b"] + base["c"]).toString(), "6")
            __check((base.size).toString(), "3")
        }
