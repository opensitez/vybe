// vybe-test: kotlin/collections/test_array_manual_copy_is_independent
// origin: languages/kotlin/tests/kotlin/test_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val source = arrayOf(1, 2, 3)
            val copy = Array(source.size) { index -> source[index] }
            copy[0] = 9
            __check((source[0]).toString(), "1")
            __check((copy[0]).toString(), "9")
        }
