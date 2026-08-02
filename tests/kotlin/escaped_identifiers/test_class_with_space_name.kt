// vybe-test: kotlin/escaped_identifiers/test_class_with_space_name
// origin: languages/kotlin/tests/kotlin/test_escaped_identifiers.rs

class `My Class` { val value = 1 }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((`My Class`().value).toString(), "1") }
