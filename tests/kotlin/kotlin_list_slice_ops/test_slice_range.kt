// vybe-test: kotlin/kotlin_list_slice_ops/test_slice_range
// origin: languages/kotlin/tests/kotlin/test_kotlin_list_slice_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = listOf(1, 2, 3, 4)
            __check((a.slice(1 until 3).toString()).toString(), "[2, 3]")
        }
