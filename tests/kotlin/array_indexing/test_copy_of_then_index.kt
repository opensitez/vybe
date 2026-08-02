// vybe-test: kotlin/array_indexing/test_copy_of_then_index
// origin: languages/kotlin/tests/kotlin/test_array_indexing.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val a = intArrayOf(9,8,7)
val b = a.copyOf(5)
b[4] = 1
__check((b.size + b[4]).toString(), "6") }
