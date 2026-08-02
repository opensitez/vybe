// vybe-test: kotlin/inner_classes/test_inner_class_with_extension_style_api
// origin: languages/kotlin/tests/kotlin/test_inner_classes.rs

class Session {
            val prefix = "S"
            inner class Formatter {
                fun apply(v: String): String = prefix + v
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val out = Session().Formatter().apply("tep")
            __check((out).toString(), "Step")
        }
