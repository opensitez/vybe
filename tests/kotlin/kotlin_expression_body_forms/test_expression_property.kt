// vybe-test: kotlin/kotlin_expression_body_forms/test_expression_property
// origin: languages/kotlin/tests/kotlin/test_kotlin_expression_body_forms.rs

class Box {
            val value: Int get() = 1 + 2
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Box().value).toString(), "3")
        }
