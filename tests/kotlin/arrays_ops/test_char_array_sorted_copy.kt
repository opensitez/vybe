// vybe-test: kotlin/arrays_ops/test_char_array_sorted_copy
// origin: languages/kotlin/tests/kotlin/test_arrays_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val chars = charArrayOf('c', 'a', 'b')
            val sorted = chars.sortedArray()
            __check((sorted.joinToString(",")).toString(), "a,b,c")
        }
