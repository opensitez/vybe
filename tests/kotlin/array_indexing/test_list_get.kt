// vybe-test: kotlin/array_indexing/test_list_get
// origin: languages/kotlin/tests/kotlin/test_array_indexing.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val a = listOf(9, 8, 7)
__check((a[2]).toString(), "7") }
