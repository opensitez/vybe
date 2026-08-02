// vybe-test: kotlin/string_templates/test_template_function_call
// origin: languages/kotlin/tests/kotlin/test_string_templates.rs

fun fmt(v: Int): String = "hash" + v.toString()
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(("v=${fmt(4)}").toString(), "v=hash4")
        }
