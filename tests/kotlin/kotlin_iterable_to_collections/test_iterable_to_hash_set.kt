// vybe-test: kotlin/kotlin_iterable_to_collections/test_iterable_to_hash_set
// origin: languages/kotlin/tests/kotlin/test_kotlin_iterable_to_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val s = listOf(1, 2, 1, 3).toHashSet()
            __check((s.contains(2).toString()).toString(), "true")
            __check((s.size).toString(), "3")
        }
