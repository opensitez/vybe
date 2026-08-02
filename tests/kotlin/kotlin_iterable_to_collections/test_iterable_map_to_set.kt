// vybe-test: kotlin/kotlin_iterable_to_collections/test_iterable_map_to_set
// origin: languages/kotlin/tests/kotlin/test_kotlin_iterable_to_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val words = listOf("ab", "c", "de")
            val lengths = words.map { it.length }.toSet()
            __check((lengths.joinToString(",")).toString(), "2,1")
        }
