// vybe-test: kotlin/collections/test_array_of_chars_and_ascii
// origin: languages/kotlin/tests/kotlin/test_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val letters = arrayOf('a', 'b', 'c')
            __check((letters[0].toString()).toString(), "a")
            __check((letters[1].code).toString(), "98")
            __check((letters.size).toString(), "3")
        }
