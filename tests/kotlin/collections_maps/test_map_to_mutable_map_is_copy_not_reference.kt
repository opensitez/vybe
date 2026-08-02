// vybe-test: kotlin/collections_maps/test_map_to_mutable_map_is_copy_not_reference
// origin: languages/kotlin/tests/kotlin/test_collections_maps.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val base = mapOf("a" to 1, "b" to 2)
            val copy = base.toMutableMap()
            copy["a"] = 9
            __check((base["a"]).toString(), "1")
            __check((copy["a"]).toString(), "9")
            __check((copy.size).toString(), "2")
            __check((base.size).toString(), "2")
        }
