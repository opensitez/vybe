// vybe-test: kotlin/array_indexing/test_multidim_array_set
// origin: languages/kotlin/tests/kotlin/test_array_indexing.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val m = arrayOf(intArrayOf(1,2), intArrayOf(3,4))
m[0][1] = 9
__check((m[0][1]).toString(), "9") }
