// vybe-test: kotlin/scope/test_scope_split_between_fields_and_locals
// origin: languages/kotlin/tests/kotlin/test_scope.rs

class Probe {
            val source = "field"
            fun valueOf(input: String): String {
                val source = input
                return this.source + "-" + source
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Probe().valueOf("local")).toString(), "field-local")
        }
