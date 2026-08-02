// vybe-test: kotlin/kotlin_char_apis/test_char_titlecase
// origin: languages/kotlin/tests/kotlin/test_kotlin_char_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val c = 'a'
            __check((c.titlecase()).toString(), "A")
            val d = 'ǈ'
            __check((d.isTitleCase()).toString(), "false")
            __check(('A'.isTitleCase()).toString(), "false")
        }
