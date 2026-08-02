// vybe-test: kotlin/collections_maps_ops/test_map_plus_assign_and_minus_assign_stability
// origin: languages/kotlin/tests/kotlin/test_collections_maps_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = mutableMapOf("a" to 1, "b" to 2, "c" to 3)
            map += mapOf("d" to 4)
            map -= "b"
            __check((map.size).toString(), "3")
            __check((map["a"] + (map["c"] ?: 0) + (map["d"] ?: 0)).toString(), "8")
            __check((map["b"] ?: -1).toString(), "-1")
        }
