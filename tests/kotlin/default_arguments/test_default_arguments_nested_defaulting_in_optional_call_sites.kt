// vybe-test: kotlin/default_arguments/test_default_arguments_nested_defaulting_in_optional_call_sites
// origin: languages/kotlin/tests/kotlin/test_default_arguments.rs

fun format(a: String, b: String = "B", c: String = "C"): String = a + b + c
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((format("A")).toString(), "ABC")
            __check((format("A", c = "X")).toString(), "AXX")
            __check((format("A", "Y", "Z")).toString(), "AYZ")
        }
