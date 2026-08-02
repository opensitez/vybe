// vybe-test: kotlin/collections_maps/test_set_contains_value_after_mutation
// origin: languages/kotlin/tests/kotlin/test_collections_maps.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val ids = mutableSetOf(10, 20)
            ids.add(30)
            ids.add(20)
            ids.remove(10)
            __check((ids.contains(10)).toString(), "false")
            __check((ids.contains(30)).toString(), "true")
            __check((ids.size).toString(), "2")
        }
