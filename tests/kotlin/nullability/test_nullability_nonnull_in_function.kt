// vybe-test: kotlin/nullability/test_nullability_nonnull_in_function
// origin: languages/kotlin/tests/kotlin/test_nullability.rs

fun upper(value: String): String = value + value
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val s: String? = "abc"
__check((upper(s!!)).toString(), "abcabc") }
