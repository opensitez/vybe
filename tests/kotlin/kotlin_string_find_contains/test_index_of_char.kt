// vybe-test: kotlin/kotlin_string_find_contains/test_index_of_char
// origin: languages/kotlin/tests/kotlin/test_kotlin_string_find_contains.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val s = "abcdef"
            __check((s.indexOf('d').toString()).toString(), "3")
            __check((s.lastIndexOf('a').toString()).toString(), "0")
        }
