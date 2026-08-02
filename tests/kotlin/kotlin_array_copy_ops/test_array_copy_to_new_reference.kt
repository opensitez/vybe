// vybe-test: kotlin/kotlin_array_copy_ops/test_array_copy_to_new_reference
// origin: languages/kotlin/tests/kotlin/test_kotlin_array_copy_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = arrayOf(1, 2, 3)
            val b = a.copyOf()
            b[0] = 9
            __check((a[0]).toString(), "1")
            __check((b[0]).toString(), "9")
        }
