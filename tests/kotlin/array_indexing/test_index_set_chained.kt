// vybe-test: kotlin/array_indexing/test_index_set_chained
// origin: languages/kotlin/tests/kotlin/test_array_indexing.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val a = mutableListOf(1, 2, 3)
a[a.lastIndex] = 9
__check((a.last()).toString(), "9") }
