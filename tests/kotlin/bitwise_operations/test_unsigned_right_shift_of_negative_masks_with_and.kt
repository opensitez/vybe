// vybe-test: kotlin/bitwise_operations/test_unsigned_right_shift_of_negative_masks_with_and
// origin: languages/kotlin/tests/kotlin/test_bitwise_operations.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val signed = -8
            val unsigned = signed ushr 2
            __check((unsigned).toString(), "1073741822")
            __check((unsigned and 0x3FFFFFFF).toString(), "2")
        }
