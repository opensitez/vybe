// vybe-test: kotlin/array_indexing/test_array_contains_indexed_access
// origin: languages/kotlin/tests/kotlin/test_array_indexing.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val a = arrayOf("x", "y", "z")
__check((a.indices.contains(1)).toString(), "true") }
