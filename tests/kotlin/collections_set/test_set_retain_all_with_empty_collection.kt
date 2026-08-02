// vybe-test: kotlin/collections_set/test_set_retain_all_with_empty_collection
// origin: languages/kotlin/tests/kotlin/test_collections_set.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mutableSetOf(1, 2, 3)
            __check((values.retainAll(emptySet<Int>())).toString(), "true")
            __check((values.isEmpty()).toString(), "true")
            __check((values.size).toString(), "0")
        }
