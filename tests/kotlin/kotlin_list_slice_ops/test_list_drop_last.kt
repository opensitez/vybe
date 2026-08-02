// vybe-test: kotlin/kotlin_list_slice_ops/test_list_drop_last
// origin: languages/kotlin/tests/kotlin/test_kotlin_list_slice_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = listOf(1, 2, 3)
            __check((a.dropLast(1).toString()).toString(), "[1, 2]")
        }
