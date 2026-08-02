// vybe-test: kotlin/builtins/test_math_pipeline_with_classification_checks
// origin: languages/kotlin/tests/kotlin/test_builtins.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = (pow(9.0, 2.0) - abs(-40.0))
            __check((value).toString(), "41")
            __check((value.isNaN()).toString(), "false")
            __check((round(sqrt(value) * 1000.0)).toString(), "6403")
        }
