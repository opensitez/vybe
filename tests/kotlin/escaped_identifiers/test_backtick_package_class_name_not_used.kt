// vybe-test: kotlin/escaped_identifiers/test_backtick_package_class_name_not_used
// origin: languages/kotlin/tests/kotlin/test_escaped_identifiers.rs

class `X-Class` { fun value() = 9 }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((`X-Class`().value()).toString(), "9") }
