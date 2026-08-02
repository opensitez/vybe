// vybe-test: kotlin/collections_maps/test_set_uniqueness_and_lookup
// origin: languages/kotlin/tests/kotlin/test_collections_maps.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = setOf(1, 2, 2, 3, 1)
            __check((values.size).toString(), "3")
            __check((values.contains(2)).toString(), "true")
            __check((values.contains(4)).toString(), "false")
        }
