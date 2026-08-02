// vybe-test: kotlin/array_indexing/test_array_subscript_after_math
// origin: languages/kotlin/tests/kotlin/test_array_indexing.rs

fun pick(values: IntArray, index: Int): Int = values[index]
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val a = intArrayOf(10, 20, 30)
__check((pick(a, 0 + 2)).toString(), "30") }
