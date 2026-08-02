// vybe-test: kotlin/collections_maps/test_map_with_null_values
// origin: languages/kotlin/tests/kotlin/test_collections_maps.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mapOf("x" to null, "y" to 2)
            __check((values["x"]).toString(), "null")
            __check((values.containsKey("x")).toString(), "true")
            __check((values["z"] ?: -1).toString(), "-1")
        }
