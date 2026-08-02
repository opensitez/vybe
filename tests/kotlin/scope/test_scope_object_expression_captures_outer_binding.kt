// vybe-test: kotlin/scope/test_scope_object_expression_captures_outer_binding
// origin: languages/kotlin/tests/kotlin/test_scope.rs

open class Base(val label: String)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var prefix = "one"
            val instance = object : Base("base") {
                val captured = prefix
            }
            __check((instance.captured).toString(), "one")
            prefix = "two"
            __check((instance.label).toString(), "base")
        }
