// vybe-test: kotlin/kotlin_numeric_boundary_checks/test_double_is_infinite_roundtrip
// origin: languages/kotlin/tests/kotlin/test_kotlin_numeric_boundary_checks.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val inf = Double.POSITIVE_INFINITY
            __check((inf.isInfinite()).toString(), "true")
            val back = inf / 2
            __check((back.isInfinite()).toString(), "true")
        }
