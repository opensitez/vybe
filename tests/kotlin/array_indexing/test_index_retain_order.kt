// vybe-test: kotlin/array_indexing/test_index_retain_order
// origin: languages/kotlin/tests/kotlin/test_array_indexing.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val a = mutableListOf(3, 2, 1)
val b = a[0] + a[2]
__check((b).toString(), "4") }
