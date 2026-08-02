// vybe-test: kotlin/collections_maps/test_set_remove_and_is_empty
// origin: languages/kotlin/tests/kotlin/test_collections_maps.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val ids = mutableSetOf(1, 2, 3)
            __check((ids.remove(2)).toString(), "true")
            __check((ids.remove(4)).toString(), "false")
            __check((ids.isNotEmpty()).toString(), "true")
            ids.remove(1)
            ids.remove(3)
            __check((ids.isEmpty()).toString(), "true")
            __check((ids.size).toString(), "0")
        }
