// vybe-test: kotlin/kotlin_iterable_to_collections/test_iterable_flatten_map
// origin: languages/kotlin/tests/kotlin/test_kotlin_iterable_to_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val outer = listOf(listOf(1, 2), listOf(2, 3), listOf(3))
            val flat = outer.flatten()
            val unique = flat.toSet()
            __check((flat.joinToString(",")).toString(), "1,2,2,3,3")
            __check((unique.joinToString(",")).toString(), "1,2,3")
        }
