// vybe-test: kotlin/escaped_identifiers/test_backtick_in_object_function
// origin: languages/kotlin/tests/kotlin/test_escaped_identifiers.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
        val o = object {
            fun `compute`(x: Int): Int = x / 2
        }
        __check((o.`compute`(8)).toString(), "4")
    }
