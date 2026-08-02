// vybe-test: kotlin/collections_maps/test_mutable_list_clear_and_reuse
// origin: languages/kotlin/tests/kotlin/test_collections_maps.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mutableListOf(1, 2, 3)
            values.clear()
            __check((values.size).toString(), "0")
            values.add(4)
            values.add(5)
            __check((values.size).toString(), "2")
            __check((values[0] + values[1]).toString(), "9")
        }
