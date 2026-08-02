// vybe-test: kotlin/collections_maps_ops/test_map_with_default_does_not_mutate_source
// origin: languages/kotlin/tests/kotlin/test_collections_maps_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val source = mapOf("one" to 1).withDefault { 99 }
            __check((source.getValue("one")).toString(), "1")
            __check((source.getValue("two")).toString(), "99")
            __check((source.toMap().containsKey("two")).toString(), "false")
        }
