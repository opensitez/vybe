// vybe-test: kotlin/kotlin_expression_body_forms/test_expression_function
// origin: languages/kotlin/tests/kotlin/test_kotlin_expression_body_forms.rs

fun double(v: Int): Int = v * 2
        fun label(v: Int) = "v" + v.toString()

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((double(3)).toString(), "6")
            __check((label(4)).toString(), "v4")
        }
