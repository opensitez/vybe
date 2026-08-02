// vybe-test: kotlin/scope/test_scope_receiver_and_outer_this_with_with_expression
// origin: languages/kotlin/tests/kotlin/test_scope.rs

class Holder {
            val source = "outer"
            fun label(tag: String): String {
                return with(this) {
                    val source = tag
                    source + "-" + this@Holder.source
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
            __check((Holder().label("inner")).toString(), "inner-outer")
        }
