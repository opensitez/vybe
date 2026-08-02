// vybe-test: kotlin/string_builtins/test_string_splitter_with_limit
// origin: languages/kotlin/tests/kotlin/test_string_builtins.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val text = "a|b|c|d"
            val pieces = text.split("|", limit = 2)
            __check((pieces.size).toString(), "2")
            __check((pieces[0]).toString(), "a")
            __check((pieces[1]).toString(), "b|c|d")
        }
