// vybe-test: kotlin/scope/test_scope_qualified_this_preserves_outer_reference
// origin: languages/kotlin/tests/kotlin/test_scope.rs

class Container {
            val factor = 3

            fun makeTag(input: String): String {
                return with(this) {
                    val factor = 10
                    input + "-" + this@Container.factor
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
            val c = Container()
            __check((c.makeTag("id")).toString(), "id-3")
        }
