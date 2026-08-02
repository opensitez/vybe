// vybe-test: kotlin/collections_set/test_set_union_union_operator
// origin: languages/kotlin/tests/kotlin/test_collections_set.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = setOf(1, 2, 3)
            val b = setOf(3, 4, 5)
            val merged = a union b
            __check((merged.size).toString(), "5")
            __check((merged.contains(4)).toString(), "true")
            __check((merged.contains(1)).toString(), "true")
        }
