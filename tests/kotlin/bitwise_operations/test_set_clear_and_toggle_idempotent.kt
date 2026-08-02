// vybe-test: kotlin/bitwise_operations/test_set_clear_and_toggle_idempotent
// origin: languages/kotlin/tests/kotlin/test_bitwise_operations.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val base = 0b1001
            val set2 = base or (1 shl 1)
            val clear2 = set2 and (1 shl 1).inv()
            val toggle = base xor (1 shl 2)
            val toggledBack = toggle xor (1 shl 2)
            __check((set2).toString(), "11")
            __check((clear2).toString(), "9")
            __check((toggle).toString(), "13")
            __check((toggledBack).toString(), "9")
        }
