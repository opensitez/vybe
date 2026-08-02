// vybe-test: kotlin/object_expressions/test_object_expression_capture_outer_mutable_var
// origin: languages/kotlin/tests/kotlin/test_object_expressions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var prefix = "left"
            val obj = object {
                fun build(value: String): String = prefix + ":" + value
            }
            __check((obj.build("a")).toString(), "left:a")
            prefix = "right"
            __check((obj.build("b")).toString(), "right:b")
        }
