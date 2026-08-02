// vybe-test: kotlin/kotlin_destructuring_maps/test_entry_iteration_order_preserved_by_map_type
// origin: languages/kotlin/tests/kotlin/test_kotlin_destructuring_maps.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = linkedMapOf("first" to 1, "second" to 2)
            val first = values.entries.first()
            val last = values.entries.last()
            __check((first.key).toString(), "first")
            __check((last.key).toString(), "second")
        }
