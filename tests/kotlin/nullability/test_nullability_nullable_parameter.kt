// vybe-test: kotlin/nullability/test_nullability_nullable_parameter
// origin: languages/kotlin/tests/kotlin/test_nullability.rs

fun printIfNotNull(v: String?): Int { return if (v == null) 0 else v.length }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((printIfNotNull(null)).toString(), "0")
__check((printIfNotNull("abc")).toString(), "3") }
