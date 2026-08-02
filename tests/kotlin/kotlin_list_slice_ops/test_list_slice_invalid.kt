// vybe-test: kotlin/kotlin_list_slice_ops/test_list_slice_invalid
// origin: languages/kotlin/tests/kotlin/test_kotlin_list_slice_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = listOf(1, 2, 3)
            val empty = a.slice(0..-1)
            __check((empty.toString()).toString(), "[]")
        }
