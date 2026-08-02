// vybe-test: kotlin/array_indexing/test_array_sort_and_index
// origin: languages/kotlin/tests/kotlin/test_array_indexing.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val a = intArrayOf(4,1,3,2)
java.util.Arrays.sort(a)
__check((a[0] + a[3]).toString(), "5") }
