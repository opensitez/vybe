// vybe-test: kotlin/type_casts/test_as_in_typed_function
// origin: languages/kotlin/tests/kotlin/test_type_casts.rs

fun getText(any: Any): String { return any as String }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((getText("ok")).toString(), "ok") }
