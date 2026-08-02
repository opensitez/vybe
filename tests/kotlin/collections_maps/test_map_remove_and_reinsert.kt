// vybe-test: kotlin/collections_maps/test_map_remove_and_reinsert
// origin: languages/kotlin/tests/kotlin/test_collections_maps.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val items = mutableMapOf("a" to 1, "b" to 2, "c" to 3)
            __check((items.remove("b")).toString(), "2")
            __check((items.remove("x")).toString(), "null")
            __check((items.size).toString(), "2")
            items["b"] = 9
            __check((items["b"]).toString(), "9")
            __check((items.size).toString(), "3")
        }
