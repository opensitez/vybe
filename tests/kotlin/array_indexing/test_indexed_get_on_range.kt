// vybe-test: kotlin/array_indexing/test_indexed_get_on_range
// origin: languages/kotlin/tests/kotlin/test_array_indexing.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val r = intArrayOf(5,6,7)
__check((r[1]).toString(), "6") }
