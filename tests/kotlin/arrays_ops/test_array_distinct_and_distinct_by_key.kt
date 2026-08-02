// vybe-test: kotlin/arrays_ops/test_array_distinct_and_distinct_by_key
// origin: languages/kotlin/tests/kotlin/test_arrays_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val items = arrayOf("aa", "ab", "b", "cc")
            val distinctByLen = items.distinctBy { it.length }
            __check((distinctByLen.joinToString(",")).toString(), "aa,b")
        }
