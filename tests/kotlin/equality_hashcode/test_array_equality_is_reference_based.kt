// vybe-test: kotlin/equality_hashcode/test_array_equality_is_reference_based
// origin: languages/kotlin/tests/kotlin/test_equality_hashcode.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val left = arrayOf(1, 2)
            val right = arrayOf(1, 2)
            __check((left == right).toString(), "false")
            __check((left === right).toString(), "false")
        }
