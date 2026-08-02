// vybe-test: kotlin/object_expressions/test_object_expression_and_outer_this_in_member_scope
// origin: languages/kotlin/tests/kotlin/test_object_expressions.rs

class Envelope(val marker: String) {
            fun make(): String {
                val obj = object {
                    fun value(): String = this@Envelope.marker
                }
                return obj.value()
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Envelope("ok").make()).toString(), "ok")
        }
