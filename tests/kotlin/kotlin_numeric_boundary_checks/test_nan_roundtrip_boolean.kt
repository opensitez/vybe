// vybe-test: kotlin/kotlin_numeric_boundary_checks/test_nan_roundtrip_boolean
// origin: languages/kotlin/tests/kotlin/test_kotlin_numeric_boundary_checks.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nan = Double.NaN
            __check((nan.isNaN()).toString(), "true")
            __check(((nan == nan)).toString(), "false")
        }
