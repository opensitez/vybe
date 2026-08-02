// vybe-test: kotlin/bitwise_operations/test_bitwise_power_of_two_checks
// origin: languages/kotlin/tests/kotlin/test_bitwise_operations.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val one = 1
            val two = 1 shl 1
            val three = 1 shl 2
            val eight = 1 shl 3
            __check((two and two).toString(), "2")
            __check((three and one).toString(), "0")
            __check((eight and 4).toString(), "0")
            __check((eight and eight).toString(), "8")
        }
