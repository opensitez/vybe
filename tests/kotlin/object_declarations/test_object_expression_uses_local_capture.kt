// vybe-test: kotlin/object_declarations/test_object_expression_uses_local_capture
// origin: languages/kotlin/tests/kotlin/test_object_declarations.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val prefix = "ok"
            val value = object {
                fun label(value: Int): String = prefix + value.toString()
            }
            __check((value.label(3)).toString(), "ok3")
        }
