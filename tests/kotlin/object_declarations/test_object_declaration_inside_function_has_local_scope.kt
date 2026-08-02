// vybe-test: kotlin/object_declarations/test_object_declaration_inside_function_has_local_scope
// origin: languages/kotlin/tests/kotlin/test_object_declarations.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            object Local {
                val value = "local"
            }

            __check((Local.value).toString(), "local")
        }
