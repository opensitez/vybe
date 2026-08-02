// vybe-test: kotlin/collections_set/test_to_set_from_list_preserves_uniqueness
// origin: languages/kotlin/tests/kotlin/test_collections_set.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf(1, 2, 2, 3, 3, 3)
            val unique = values.toSet()
            __check((unique.size).toString(), "3")
            __check((unique.contains(3)).toString(), "true")
            __check((unique.toString()).toString(), "[1, 2, 3]")
        }
