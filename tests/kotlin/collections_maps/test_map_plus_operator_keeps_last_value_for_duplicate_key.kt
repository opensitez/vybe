// vybe-test: kotlin/collections_maps/test_map_plus_operator_keeps_last_value_for_duplicate_key
// origin: languages/kotlin/tests/kotlin/test_collections_maps.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val left = mapOf("a" to 1, "b" to 2)
            val right = mapOf("b" to 9, "c" to 3)
            val merged = left + right
            __check((merged.size).toString(), "3")
            __check((merged["b"]).toString(), "9")
            __check((merged["c"]).toString(), "3")
        }
