// vybe-test: kotlin/collections_maps/test_map_duplicate_keys_keep_the_last_value
// origin: languages/kotlin/tests/kotlin/test_collections_maps.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val scores = mapOf("a" to 1, "b" to 2, "a" to 7, "b" to 9)
            __check((scores.size).toString(), "2")
            __check((scores["a"]).toString(), "7")
            __check((scores["b"]).toString(), "9")
        }
