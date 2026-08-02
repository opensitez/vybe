// vybe-test: kotlin/kotlin_array_copy_ops/test_array_copy_of_range
// origin: languages/kotlin/tests/kotlin/test_kotlin_array_copy_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = arrayOf("a", "b", "c", "d")
            val b = a.copyOfRange(1, 3)
            __check((b.toString()).toString(), "[b, c]")
        }
