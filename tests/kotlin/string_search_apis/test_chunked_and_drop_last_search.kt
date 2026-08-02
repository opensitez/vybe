// vybe-test: kotlin/string_search_apis/test_chunked_and_drop_last_search
// origin: languages/kotlin/tests/kotlin/test_string_search_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val source = "abcdef"
            __check((source.chunked(2).joinToString("|")).toString(), "ab|cd|ef")
            __check((source.takeLast(3)).toString(), "def")
            __check((source.dropLast(2)).toString(), "abcd")
        }
