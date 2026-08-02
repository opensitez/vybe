// vybe-test: kotlin/collections_set/test_set_union_operator_plus
// origin: languages/kotlin/tests/kotlin/test_collections_set.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = setOf(1, 2)
            val b = setOf(2, 3)
            val merged = a + b
            __check((merged.size).toString(), "3")
            __check((merged.contains(3)).toString(), "true")
            __check(((a + b).contains(1)).toString(), "true")
        }
