// vybe-test: kotlin/collections_maps/test_mutable_list_add_remove
// origin: languages/kotlin/tests/kotlin/test_collections_maps.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mutableListOf(1, 3)
            values.add(5)
            values.removeAt(1)
            __check((values.size).toString(), "2")
            __check((values[1]).toString(), "5")
        }
