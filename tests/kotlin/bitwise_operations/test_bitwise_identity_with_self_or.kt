// vybe-test: kotlin/bitwise_operations/test_bitwise_identity_with_self_or
// origin: languages/kotlin/tests/kotlin/test_bitwise_operations.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf(0, 1, 2, 3, 255)
            val same = values.map { it or it }
            val plus = values.map { it or 0 }
            __check((same.joinToString(",")).toString(), "0,1,2,3,255")
            __check((plus.joinToString(",")).toString(), "0,1,2,3,255")
        }
