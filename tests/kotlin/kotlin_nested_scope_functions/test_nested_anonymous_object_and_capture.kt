// vybe-test: kotlin/kotlin_nested_scope_functions/test_nested_anonymous_object_and_capture
// origin: languages/kotlin/tests/kotlin/test_kotlin_nested_scope_functions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            fun make(prefix: String) = object {
                fun label(v: Int) = prefix + v
            }

            val p = make("x")
            __check((p.label(9)).toString(), "x9")
        }
