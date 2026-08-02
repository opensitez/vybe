// vybe-test: kotlin/escaped_identifiers/test_backtick_in_generic_type_name
// origin: languages/kotlin/tests/kotlin/test_escaped_identifiers.rs

typealias `Alias Name` = Map<String, Int>
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
    val x: `Alias Name` = mapOf("a" to 1)
    __check((x["a"]).toString(), "1")
}
