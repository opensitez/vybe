// vybe-test: kotlin/kotlin_char_apis/test_char_sequence_navigation
// origin: languages/kotlin/tests/kotlin/test_kotlin_char_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val s = "Abc"
            __check((s.first()).toString(), "A")
            __check((s.last()).toString(), "c")
            __check((s.elementAt(1)).toString(), "b")
            __check((s[2]).toString(), "c")
        }
