// vybe-test: kotlin/default_arguments/test_default_arguments_default_value_is_literal_object
// origin: languages/kotlin/tests/kotlin/test_default_arguments.rs

fun paint(color: String = "red", opacity: Double = 1.0): String = color + ":" + opacity
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((paint()).toString(), "red:1.0")
            __check((paint(opacity = 0.5)).toString(), "red:0.5")
        }
