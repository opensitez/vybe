// vybe-test: kotlin/kotlin_resource_management/test_nested_uses_close_outer_and_inner
// origin: languages/kotlin/tests/kotlin/test_kotlin_resource_management.rs

class Token(val name: String) : AutoCloseable {
            var calls: Int = 0
            override fun close() {
                calls += 1
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = Token("a")
            val b = Token("b")
            a.use {
                b.use {
                    __check((a.calls + b.calls).toString(), "0")
                }
            }
            __check((a.calls).toString(), "1")
            __check((b.calls).toString(), "1")
        }
