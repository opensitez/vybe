// vybe-test: kotlin/bitwise_operations/test_bitwise_identity_with_self_xor
// origin: languages/kotlin/tests/kotlin/test_bitwise_operations.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf(0, 1, 2, 3, 255)
            val unchanged = values.map { it xor it }
            val back = values.map { (it xor 0) xor it }
            __check((unchanged.joinToString(",")).toString(), "0,0,0,0,0")
            __check((back.joinToString(",")).toString(), "0,1,2,3,255")
        }
