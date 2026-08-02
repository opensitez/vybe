// vybe-test: kotlin/nested_classes/test_nested_class_with_function_calls
// origin: languages/kotlin/tests/kotlin/test_nested_classes.rs

class Registry {
            class Entry(val name: String)

            companion object {
                fun make(name: String): Entry = Entry(name)
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val e = Registry.make("x")
            __check((e.name).toString(), "x")
        }
