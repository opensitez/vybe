// vybe-test: kotlin/array_indexing/test_boolean_array_set
// origin: languages/kotlin/tests/kotlin/test_array_indexing.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val a = booleanArrayOf(false, true)
a[0] = true
__check((a[0]).toString(), "true") }
