// vybe-test: kotlin/array_indexing/test_string_char_at
// origin: languages/kotlin/tests/kotlin/test_array_indexing.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val s = "kotlin"
__check((s[1]).toString(), "o") }
