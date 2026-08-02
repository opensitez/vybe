// vybe-test: kotlin/array_indexing/test_char_array_slice_to_string
// origin: languages/kotlin/tests/kotlin/test_array_indexing.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val a = charArrayOf('x','y','z')
__check((a.joinToString("")).toString(), "xyz") }
