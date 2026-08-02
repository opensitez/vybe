// vybe-test: kotlin/kotlin_char_array_apis/test_char_array_empty
// origin: languages/kotlin/tests/kotlin/test_kotlin_char_array_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val e = charArrayOf()
            __check((e.size).toString(), "0")
            __check((e.isEmpty().toString()).toString(), "true")
        }
