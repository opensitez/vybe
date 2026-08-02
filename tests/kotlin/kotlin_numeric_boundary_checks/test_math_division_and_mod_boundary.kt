// vybe-test: kotlin/kotlin_numeric_boundary_checks/test_math_division_and_mod_boundary
// origin: languages/kotlin/tests/kotlin/test_kotlin_numeric_boundary_checks.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((0 / Int.MAX_VALUE).toString(), "0")
            __check((Int.MIN_VALUE % -1).toString(), "0")
            __check((7 % 3).toString(), "1")
        }
