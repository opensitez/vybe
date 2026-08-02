// vybe-test: kotlin/extension_functions/test_extension_function_on_class_instance
// origin: languages/kotlin/tests/kotlin/test_extension_functions.rs

class Box(val value: Int)

        fun Box.labeled(prefix: String): String = prefix + ":" + value

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Box(7).labeled("v")).toString(), "v:7")
        }
