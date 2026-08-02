// vybe-test: kotlin/kotlin_iterable_to_collections/test_to_typed_array
// origin: languages/kotlin/tests/kotlin/test_kotlin_iterable_to_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val ints = listOf(1, 2, 3).toIntArray()
            __check((ints.joinToString(",")).toString(), "1,2,3")
            val chars = listOf('a', 'b').toCharArray()
            __check((chars.joinToString(",")).toString(), "a,b")
        }
