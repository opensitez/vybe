// vybe-test: kotlin/collections_maps/test_map_size_after_clear
// origin: languages/kotlin/tests/kotlin/test_collections_maps.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val data = mutableMapOf("a" to 1, "b" to 2)
            data.clear()
            __check((data.isEmpty()).toString(), "true")
            data["z"] = 8
            __check((data.size).toString(), "1")
            __check((data["z"]).toString(), "8")
        }
