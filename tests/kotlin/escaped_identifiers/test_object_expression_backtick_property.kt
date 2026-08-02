// vybe-test: kotlin/escaped_identifiers/test_object_expression_backtick_property
// origin: languages/kotlin/tests/kotlin/test_escaped_identifiers.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val o = object { val `x y` = 3 }
            __check((o.`x y`).toString(), "3")
        }
