// vybe-test: kotlin/array_indexing/test_last_index
// origin: languages/kotlin/tests/kotlin/test_array_indexing.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val a = intArrayOf(4, 5, 6)
__check((a[a.lastIndex]).toString(), "6") }
