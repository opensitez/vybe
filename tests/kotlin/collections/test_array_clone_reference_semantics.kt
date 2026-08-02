// vybe-test: kotlin/collections/test_array_clone_reference_semantics
// origin: languages/kotlin/tests/kotlin/test_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val original = arrayOf(1, 2, 3)
            val shared = original
            shared[1] = 99
            __check((original[1]).toString(), "99")
            __check((shared[1]).toString(), "99")
        }
