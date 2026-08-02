// vybe-test: kotlin/collections_maps_ops/test_map_get_or_put_without_recomputing_for_present_key
// origin: languages/kotlin/tests/kotlin/test_collections_maps_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var computed = 0
            val map = mutableMapOf("present" to 1)
            val value1 = map.getOrPut("present") {
                computed += 1
                99
            }
            val value2 = map.getOrPut("missing") {
                computed += 1
                77
            }
            __check((value1).toString(), "1")
            __check((value2).toString(), "77")
            __check((computed).toString(), "1")
        }
