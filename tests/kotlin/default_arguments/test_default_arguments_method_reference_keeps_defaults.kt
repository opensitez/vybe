// vybe-test: kotlin/default_arguments/test_default_arguments_method_reference_keeps_defaults
// origin: languages/kotlin/tests/kotlin/test_default_arguments.rs

fun decorate(text: String, marker: String = "*"): String = marker + text + marker
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val f = ::decorate
            __check((f("x")).toString(), "*x*")
            __check((f("y", "#")).toString(), "#y#")
        }
