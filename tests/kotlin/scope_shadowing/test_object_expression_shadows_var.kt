// vybe-test: kotlin/scope_shadowing/test_object_expression_shadows_var
// origin: languages/kotlin/tests/kotlin/test_scope_shadowing.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var tag = "outer"
            val obj = object {
                val tag = "inner"
                fun value(): String = tag
            }
            __check((obj.value()).toString(), "inner")
            __check((tag).toString(), "outer")
        }
