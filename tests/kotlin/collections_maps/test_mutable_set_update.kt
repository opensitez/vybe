// vybe-test: kotlin/collections_maps/test_mutable_set_update
// origin: languages/kotlin/tests/kotlin/test_collections_maps.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mutableSetOf(1, 2)
            values.add(2)
            values.add(3)
            __check((values.size).toString(), "3")
            values.remove(1)
            __check((values.contains(1)).toString(), "false")
        }
