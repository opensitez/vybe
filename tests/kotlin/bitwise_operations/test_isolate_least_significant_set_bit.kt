// vybe-test: kotlin/bitwise_operations/test_isolate_least_significant_set_bit
// origin: languages/kotlin/tests/kotlin/test_bitwise_operations.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = 0b1011000
            val lsb = value and (-value)
            __check((lsb).toString(), "8")
            __check(((value and (value - 1))).toString(), "88")
        }
