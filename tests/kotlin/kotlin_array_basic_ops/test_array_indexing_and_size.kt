// vybe-test: kotlin/kotlin_array_basic_ops/test_array_indexing_and_size
// origin: languages/kotlin/tests/kotlin/test_kotlin_array_basic_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = arrayOf(10, 20, 30)
            __check((a.size).toString(), "3")
            __check((a[1] + a[2]).toString(), "50")
        }
