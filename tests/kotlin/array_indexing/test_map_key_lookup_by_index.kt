// vybe-test: kotlin/array_indexing/test_map_key_lookup_by_index
// origin: languages/kotlin/tests/kotlin/test_array_indexing.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val a = mapOf(1 to "a", 2 to "b")
__check((a[2]).toString(), "b") }
