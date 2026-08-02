// vybe-test: kotlin/companion_objects/test_named_companion_object_can_be_used_as_type
// origin: languages/kotlin/tests/kotlin/test_companion_objects.rs

class Parser {
            companion object Validator {
                fun ok(value: String): Boolean = value.isNotEmpty()
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val valid = Parser.Validator.ok("x")
            val invalid = Parser.Validator.ok("")
            __check((valid).toString(), "true")
            __check((invalid).toString(), "false")
        }
