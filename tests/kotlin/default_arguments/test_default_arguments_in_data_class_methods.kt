// vybe-test: kotlin/default_arguments/test_default_arguments_in_data_class_methods
// origin: languages/kotlin/tests/kotlin/test_default_arguments.rs

data class Box(val value: Int, val label: String = "x")
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = Box(1)
            val b = a.copy(label = "y")
            __check((a.label).toString(), "x")
            __check((b.value).toString(), "1")
            __check((b.label).toString(), "y")
        }
