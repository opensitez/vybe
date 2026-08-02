// vybe-test: kotlin/nullability/test_nullability_nested_optional
// origin: languages/kotlin/tests/kotlin/test_nullability.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val a: String? = null
val b = a ?: ("z")
val c = b + "oo"
__check((c).toString(), "zoo") }
