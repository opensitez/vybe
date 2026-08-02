// vybe-test: kotlin/default_arguments/test_default_arguments_nested_defaults_with_defaults_on_named_type
// origin: languages/kotlin/tests/kotlin/test_default_arguments.rs

fun line(a: String = "a", b: String = a): String = a + b
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((line()).toString(), "aa")
            __check((line("x")).toString(), "xx")
        }
