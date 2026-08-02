// vybe-test: kotlin/escaped_identifiers/test_extension_on_string_backtick
// origin: languages/kotlin/tests/kotlin/test_escaped_identifiers.rs

fun String.`dash`(): String = this + "!"
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check(("ok".`dash`()).toString(), "ok!") }
