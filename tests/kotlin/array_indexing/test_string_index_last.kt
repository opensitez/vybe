// vybe-test: kotlin/array_indexing/test_string_index_last
// origin: languages/kotlin/tests/kotlin/test_array_indexing.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val s = "hello"
__check((s[s.lastIndex]).toString(), "o") }
