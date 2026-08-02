// vybe-test: kotlin/array_indexing/test_char_array_get
// origin: languages/kotlin/tests/kotlin/test_array_indexing.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val a = charArrayOf('a', 'b', 'c')
__check((a[0]).toString(), "a") }
