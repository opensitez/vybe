// vybe-test: kotlin/nullability/test_nullability_or_return
// origin: languages/kotlin/tests/kotlin/test_nullability.rs

fun ensure(v: String?): String { return v ?: "empty" }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((ensure(null)).toString(), "empty")
__check((ensure("ok")).toString(), "ok") }
