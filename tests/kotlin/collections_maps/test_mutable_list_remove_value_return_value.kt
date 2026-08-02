// vybe-test: kotlin/collections_maps/test_mutable_list_remove_value_return_value
// origin: languages/kotlin/tests/kotlin/test_collections_maps.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mutableListOf(1, 2, 3)
            __check((values.remove(2)).toString(), "true")
            __check((values.remove(9)).toString(), "false")
            __check((values.size).toString(), "2")
        }
