// vybe-test: kotlin/kotlin_numeric_boundary_checks/test_unsigned_conversion_roundtrip
// origin: languages/kotlin/tests/kotlin/test_kotlin_numeric_boundary_checks.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val n: Int = -1
            val u = n.toUInt()
            __check((u).toString(), "4294967295")
            __check((u.toInt()).toString(), "-1")
        }
