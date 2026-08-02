// vybe-test: kotlin/bitwise_operations/test_bitwise_identity_with_self_and
// origin: languages/kotlin/tests/kotlin/test_bitwise_operations.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf(0, 1, 2, 3, 255)
            val kept = values.map { it and it }
            val zeroed = values.map { it and 0 }
            __check((kept.joinToString(",")).toString(), "0,1,2,3,255")
            __check((zeroed.joinToString(",")).toString(), "0,0,0,0,0")
        }
