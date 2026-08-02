// vybe-test: kotlin/object_declarations/test_nested_object_declaration
// origin: languages/kotlin/tests/kotlin/test_object_declarations.rs

class Holder {
            object Defaults {
                val label = "ok"
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Holder.Defaults.label).toString(), "ok")
        }
