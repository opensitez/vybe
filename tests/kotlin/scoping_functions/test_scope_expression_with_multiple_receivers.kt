// vybe-test: kotlin/scoping_functions/test_scope_expression_with_multiple_receivers
// origin: languages/kotlin/tests/kotlin/test_scoping_functions.rs

class Holder {
            fun make(prefix: String): String = with(this) { "$prefix:value" }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = Holder().run {
                make("start").also { }
            }
            __check((value).toString(), "start:value")
        }
