// vybe-test: kotlin/type_casts/test_boolean_as_cast_in_function
// origin: languages/kotlin/tests/kotlin/test_type_casts.rs

fun isTrue(value: Any?): Boolean { return value as? Boolean ?: false }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((isTrue(true)).toString(), "true")
__check((isTrue("n")).toString(), "false") }
