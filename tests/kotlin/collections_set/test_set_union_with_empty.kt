// vybe-test: kotlin/collections_set/test_set_union_with_empty
// origin: languages/kotlin/tests/kotlin/test_collections_set.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = setOf(1, 2)
            val merged = values + emptySet<Int>()
            __check((merged.size).toString(), "2")
            __check((merged == values).toString(), "true")
            __check((emptySet<Int>() + values == values).toString(), "true")
        }
