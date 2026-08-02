// vybe-test: kotlin/collections_maps/test_list_empty_properties
// origin: languages/kotlin/tests/kotlin/test_collections_maps.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf<Int>()
            __check((values.size).toString(), "0")
            __check((values.isEmpty()).toString(), "true")
        }
