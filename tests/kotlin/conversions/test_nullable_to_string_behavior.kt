// vybe-test: kotlin/conversions/test_nullable_to_string_behavior
// origin: languages/kotlin/tests/kotlin/test_conversions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: String? = null
            __check((value == null).toString(), "true")
            val fallback = value?.toString() ?: "none"
            __check((fallback).toString(), "none")
        }
