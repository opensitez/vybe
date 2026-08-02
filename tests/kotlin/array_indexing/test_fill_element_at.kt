// vybe-test: kotlin/array_indexing/test_fill_element_at
// origin: languages/kotlin/tests/kotlin/test_array_indexing.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val a = intArrayOf(1, 2, 3)
java.util.Arrays.fill(a, 1, 2, 9)
__check((a[1]).toString(), "9") }
