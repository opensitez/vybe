// vybe-test: kotlin/collections_maps_ops/test_map_plus_assign_mutates_same_map
// origin: languages/kotlin/tests/kotlin/test_collections_maps_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = mutableMapOf("a" to 1)
            map += mapOf("b" to 2, "c" to 3)
            __check((map.size).toString(), "3")
            __check((map["b"] + map["c"]).toString(), "5")
        }
