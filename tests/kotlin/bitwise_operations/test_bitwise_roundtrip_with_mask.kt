// vybe-test: kotlin/bitwise_operations/test_bitwise_roundtrip_with_mask
// origin: languages/kotlin/tests/kotlin/test_bitwise_operations.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val original = 0b10101010
            val mask = 0b11110000
            val hidden = original and mask
            val shown = original and mask.inv()
            val visible = (original and mask.inv())
            __check((hidden).toString(), "160")
            __check((shown).toString(), "10")
            __check((visible).toString(), "10")
            __check((hidden + visible).toString(), "170")
        }
