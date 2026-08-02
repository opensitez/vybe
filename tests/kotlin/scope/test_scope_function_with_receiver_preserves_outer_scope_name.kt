// vybe-test: kotlin/scope/test_scope_function_with_receiver_preserves_outer_scope_name
// origin: languages/kotlin/tests/kotlin/test_scope.rs

class Box {
            val value = "outer"
            fun label(): String {
                return with(this) {
                    val value = "inner"
                    value + "-" + this.value
                }
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Box().label()).toString(), "inner-outer")
        }
