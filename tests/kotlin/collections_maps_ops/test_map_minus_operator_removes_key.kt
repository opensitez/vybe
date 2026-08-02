// vybe-test: kotlin/collections_maps_ops/test_map_minus_operator_removes_key
// origin: languages/kotlin/tests/kotlin/test_collections_maps_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val source = mapOf("a" to 1, "b" to 2, "c" to 3)
            val reduced = source - "b"
            __check((reduced.size).toString(), "2")
            __check((reduced.containsKey("b")).toString(), "false")
            __check((reduced["a"] + reduced["c"]).toString(), "4")
        }
