// vybe-test: kotlin/escaped_identifiers/test_backtick_in_type_parameter
// origin: languages/kotlin/tests/kotlin/test_escaped_identifiers.rs

class `Holder Type`<T>(val `value data`: T)
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
    val h = `Holder Type`(5)
    __check((h.`value data`).toString(), "5")
}
