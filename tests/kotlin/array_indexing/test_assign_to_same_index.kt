// vybe-test: kotlin/array_indexing/test_assign_to_same_index
// origin: languages/kotlin/tests/kotlin/test_array_indexing.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val a = IntArray(1)
a[0] = a[0] + 1
__check((a[0]).toString(), "1") }
