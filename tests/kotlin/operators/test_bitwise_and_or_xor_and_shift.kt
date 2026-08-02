// vybe-test: kotlin/operators/test_bitwise_and_or_xor_and_shift
// origin: languages/kotlin/tests/kotlin/test_operators.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((1 shl 4).toString(), "16")
            __check((16 shr 2).toString(), "4")
            __check((5 and 3).toString(), "1")
            __check((5 or 2).toString(), "7")
            __check((5 xor 2).toString(), "7")
            __check((5.inv()).toString(), "-6")
        }
