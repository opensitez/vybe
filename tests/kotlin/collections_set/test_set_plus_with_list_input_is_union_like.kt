// vybe-test: kotlin/collections_set/test_set_plus_with_list_input_is_union_like
// origin: languages/kotlin/tests/kotlin/test_collections_set.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val source = setOf(1, 2)
            val merged = source + listOf(2, 3, 4)
            __check((merged.size).toString(), "4")
            __check((merged.contains(3)).toString(), "true")
            __check((merged.contains(1)).toString(), "true")
        }
