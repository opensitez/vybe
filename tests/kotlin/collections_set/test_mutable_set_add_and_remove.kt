// vybe-test: kotlin/collections_set/test_mutable_set_add_and_remove
// origin: languages/kotlin/tests/kotlin/test_collections_set.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mutableSetOf(1, 2)
            __check((values.add(3)).toString(), "true")
            __check((values.remove(2)).toString(), "true")
            __check((values.size).toString(), "2")
            __check((values.contains(2)).toString(), "false")
        }
