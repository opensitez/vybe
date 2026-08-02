// vybe-test: kotlin/kotlin_string_escapes/test_backslash_escape
// origin: languages/kotlin/tests/kotlin/test_kotlin_string_escapes.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(("path:" + "c\\\\temp").toString(), "path:c\\temp")
            __check(("quote:" + "\"").toString(), "quote:\"")
        }
