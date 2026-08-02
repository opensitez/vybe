// vybe-test: kotlin/array_indexing/test_index_before_set
// origin: languages/kotlin/tests/kotlin/test_array_indexing.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val a = intArrayOf(1, 2)
a[0] = a[1] + 4
__check((a.joinToString(",")).toString(), "6,2") }
