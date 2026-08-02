// vybe-test: kotlin/object_declarations/test_object_expression_and_object_declaration_difference
// origin: languages/kotlin/tests/kotlin/test_object_declarations.rs

object Labeler {
            fun label(value: Int): String = "v" + value
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = object {
                fun label(value: Int): String = "local" + value
            }
            __check((value.label(4)).toString(), "local4")
            __check((Labeler.label(4)).toString(), "v4")
        }
