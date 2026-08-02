// vybe-test: kotlin/string_builtins/test_string_replace_and_index_of
// origin: languages/kotlin/tests/kotlin/test_string_builtins.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val text = "aa-bb-aa"
            __check((text.replace("aa", "x")).toString(), "x-bb-x")
            __check((text.indexOf("bb")).toString(), "3")
        }
