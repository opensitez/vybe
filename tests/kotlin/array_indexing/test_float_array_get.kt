// vybe-test: kotlin/array_indexing/test_float_array_get
// origin: languages/kotlin/tests/kotlin/test_array_indexing.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val a = floatArrayOf(1.5f, 2.5f)
__check((a[1]).toString(), "2.5") }
