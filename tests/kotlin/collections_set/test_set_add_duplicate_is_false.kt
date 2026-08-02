// vybe-test: kotlin/collections_set/test_set_add_duplicate_is_false
// origin: languages/kotlin/tests/kotlin/test_collections_set.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mutableSetOf(1, 2)
            __check((values.add(2)).toString(), "false")
            __check((values.size).toString(), "2")
        }
