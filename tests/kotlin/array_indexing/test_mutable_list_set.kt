// vybe-test: kotlin/array_indexing/test_mutable_list_set
// origin: languages/kotlin/tests/kotlin/test_array_indexing.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val a = mutableListOf(1, 2, 3)
a[1] = 5
__check((a[1]).toString(), "5") }
