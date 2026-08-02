// vybe-test: kotlin/array_indexing/test_long_array_reduce_index
// origin: languages/kotlin/tests/kotlin/test_array_indexing.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val a = longArrayOf(2L, 4L, 6L)
__check((a[1] + a[2]).toString(), "10") }
