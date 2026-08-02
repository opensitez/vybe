// vybe-test: kotlin/collections_maps/test_set_starts_empty_and_adds_unique
// origin: languages/kotlin/tests/kotlin/test_collections_maps.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val ids = mutableSetOf<Int>()
            ids.add(1)
            ids.add(2)
            ids.add(2)
            __check((ids.size).toString(), "2")
            __check((ids.contains(1)).toString(), "true")
            __check((ids.contains(3)).toString(), "false")
        }
