// vybe-test: kotlin/array_indexing/test_copy_from_range
// origin: languages/kotlin/tests/kotlin/test_array_indexing.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val a = intArrayOf(1, 2, 3, 4)
val b = a.copyOfRange(1, 3)
__check((b.joinToString(",")).toString(), "2,3") }
