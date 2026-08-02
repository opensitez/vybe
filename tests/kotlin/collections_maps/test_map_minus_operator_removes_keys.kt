// vybe-test: kotlin/collections_maps/test_map_minus_operator_removes_keys
// origin: languages/kotlin/tests/kotlin/test_collections_maps.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val base = mapOf("a" to 1, "b" to 2, "c" to 3)
            val narrowed = base - "b"
            __check((narrowed.size).toString(), "2")
            __check((narrowed.containsKey("b")).toString(), "false")
            __check((narrowed["c"]).toString(), "3")
        }
