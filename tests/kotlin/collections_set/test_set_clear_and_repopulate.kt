// vybe-test: kotlin/collections_set/test_set_clear_and_repopulate
// origin: languages/kotlin/tests/kotlin/test_collections_set.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mutableSetOf(1, 2)
            values.clear()
            values.add(9)
            values.add(9)
            __check((values.size).toString(), "1")
            __check((values.contains(9)).toString(), "true")
        }
