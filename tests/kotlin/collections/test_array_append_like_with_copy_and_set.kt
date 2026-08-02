// vybe-test: kotlin/collections/test_array_append_like_with_copy_and_set
// origin: languages/kotlin/tests/kotlin/test_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val base = arrayOf(1, 2)
            val extended = Array(base.size + 1) { index ->
                if (index < base.size) base[index] else 3
            }
            __check((extended.size).toString(), "3")
            __check((extended[2]).toString(), "3")
            __check((extended[0] + extended[1] + extended[2]).toString(), "6")
        }
