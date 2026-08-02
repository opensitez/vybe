// vybe-test: kotlin/array_indexing/test_fill_array_then_get
// origin: languages/kotlin/tests/kotlin/test_array_indexing.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val a = IntArray(3)
a.fill(5)
__check((a[2]).toString(), "5") }
