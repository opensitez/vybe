// vybe-test: kotlin/collections_maps/test_map_with_default_keeps_original_map_without_side_updates
// origin: languages/kotlin/tests/kotlin/test_collections_maps.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val base = mapOf("known" to 5).withDefault { 77 }
            __check((base.getValue("known")).toString(), "5")
            __check((base.getValue("missing")).toString(), "77")
            __check((base.toMap().containsKey("missing")).toString(), "false")
        }
