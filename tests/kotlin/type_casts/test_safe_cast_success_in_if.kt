// vybe-test: kotlin/type_casts/test_safe_cast_success_in_if
// origin: languages/kotlin/tests/kotlin/test_type_casts.rs

fun describe(value: Any): String { return if (value is Int) "int" else "not int" }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((describe(10)).toString(), "int")
__check((describe("x")).toString(), "not int") }
