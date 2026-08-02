// vybe-test: kotlin/collections_maps/test_map_plus_assign_adds_and_removes_keys
// origin: languages/kotlin/tests/kotlin/test_collections_maps.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val source = linkedMapOf("a" to 1)
            source += mapOf("b" to 2, "c" to 3)
            source += mapOf("c" to 4)
            __check((source["b"] + source["c"]).toString(), "6")
            __check((source.size).toString(), "3")
        }
