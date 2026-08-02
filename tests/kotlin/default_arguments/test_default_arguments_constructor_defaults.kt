// vybe-test: kotlin/default_arguments/test_default_arguments_constructor_defaults
// origin: languages/kotlin/tests/kotlin/test_default_arguments.rs

class Box(val value: Int = 1, val label: String = "x")
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = Box()
            val b = Box(2)
            val c = Box(label = "z")
            __check((a.value).toString(), "1")
            __check((a.label).toString(), "x")
            __check((b.value).toString(), "2")
            __check((c.label).toString(), "z")
        }
