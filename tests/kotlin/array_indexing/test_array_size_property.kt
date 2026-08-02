// vybe-test: kotlin/array_indexing/test_array_size_property
// origin: languages/kotlin/tests/kotlin/test_array_indexing.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val a = IntArray(4)
__check((a.size).toString(), "4") }
