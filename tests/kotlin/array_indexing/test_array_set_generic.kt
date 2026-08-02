// vybe-test: kotlin/array_indexing/test_array_set_generic
// origin: languages/kotlin/tests/kotlin/test_array_indexing.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val a = arrayOf("x", "y")
a[0] = "z"
__check((a[0]).toString(), "z") }
