// vybe-test: kotlin/kotlin_string_escapes/test_tab_escape_text
// origin: languages/kotlin/tests/kotlin/test_kotlin_string_escapes.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(("x" + "\\t" + "y").toString(), "x\\ty")
            __check(("a" + "\\r" + "b").toString(), "a\\rb")
        }
