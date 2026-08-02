// vybe-test: kotlin/nullability/test_nullability_not_null_assertion
// origin: languages/kotlin/tests/kotlin/test_nullability.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val value: String? = "k"
__check((value!!).toString(), "k") }
