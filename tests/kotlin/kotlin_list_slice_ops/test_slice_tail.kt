// vybe-test: kotlin/kotlin_list_slice_ops/test_slice_tail
// origin: languages/kotlin/tests/kotlin/test_kotlin_list_slice_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = listOf(1, 2, 3)
            __check((a.subList(1, a.size).toString()).toString(), "[2, 3]")
        }
