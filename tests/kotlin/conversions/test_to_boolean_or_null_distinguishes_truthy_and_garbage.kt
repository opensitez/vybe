// vybe-test: kotlin/conversions/test_to_boolean_or_null_distinguishes_truthy_and_garbage
// origin: languages/kotlin/tests/kotlin/test_conversions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(("true".toBooleanOrNull() ?: "null").toString(), "true")
            __check(("FALSE".toBooleanOrNull() ?: "null").toString(), "false")
            __check(("yes".toBooleanOrNull() ?: "null").toString(), "null")
            __check(("0".toBooleanOrNull() ?: "null").toString(), "null")
        }
