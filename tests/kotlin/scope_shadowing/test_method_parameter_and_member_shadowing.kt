// vybe-test: kotlin/scope_shadowing/test_method_parameter_and_member_shadowing
// origin: languages/kotlin/tests/kotlin/test_scope_shadowing.rs

class Holder {
            val value: String = "member"
            fun label(value: String): String {
                return value
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val holder = Holder()
            __check((holder.label("param")).toString(), "param")
            __check((holder.value).toString(), "member")
        }
