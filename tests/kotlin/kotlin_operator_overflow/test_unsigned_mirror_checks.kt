// vybe-test: kotlin/kotlin_operator_overflow/test_unsigned_mirror_checks
// origin: languages/kotlin/tests/kotlin/test_kotlin_operator_overflow.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((255u.toByte().toInt()).toString(), "-1")
            __check(((-1).toUInt()).toString(), "4294967295")
        }
