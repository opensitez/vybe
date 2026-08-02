// vybe-test: kotlin/array_indexing/test_set_and_get_characters
// origin: languages/kotlin/tests/kotlin/test_array_indexing.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val s = StringBuilder("abc")
s[1] = 'z'
__check((s.toString()).toString(), "azc") }
