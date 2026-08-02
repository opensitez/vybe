// vybe-test: kotlin/kotlin_list_slice_ops/test_list_slice_step
// origin: languages/kotlin/tests/kotlin/test_kotlin_list_slice_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = listOf(1, 2, 3, 4, 5)
            __check((a.slice(0..4 step 2).toString()).toString(), "[1, 3, 5]")
        }
