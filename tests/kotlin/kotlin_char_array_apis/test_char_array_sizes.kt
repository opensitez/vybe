// vybe-test: kotlin/kotlin_char_array_apis/test_char_array_sizes
// origin: languages/kotlin/tests/kotlin/test_kotlin_char_array_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val data = charArrayOf('a', 'b', 'c')
            __check((data.size).toString(), "3")
        }
