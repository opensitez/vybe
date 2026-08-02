// vybe-test: kotlin/default_arguments/test_default_arguments_in_nested_class_context
// origin: languages/kotlin/tests/kotlin/test_default_arguments.rs

class Host {
            fun outer(prefix: String = "p", suffix: String = "s") : String = prefix + suffix
            class Child {
                fun inner(tag: String = "t") : String = tag
            }
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val h = Host()
            val c = Host.Child()
            __check((h.outer()).toString(), "ps")
            __check((c.inner()).toString(), "t")
            __check((c.inner("x")).toString(), "x")
        }
