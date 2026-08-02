// vybe-test: kotlin/primitive_array_apis/test_char_array_to_string
// origin: languages/kotlin/tests/kotlin/test_primitive_array_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = charArrayOf('k', 'o', 't')
            __check((values.concatToString()).toString(), "kot")
            __check((values.size).toString(), "3")
        }
