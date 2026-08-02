// vybe-test: kotlin/kotlin_char_array_apis/test_char_array_update
// origin: languages/kotlin/tests/kotlin/test_kotlin_char_array_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val data = charArrayOf('a', 'b')
            data[1] = 'z'
            __check((data[1].toString()).toString(), "z")
        }
