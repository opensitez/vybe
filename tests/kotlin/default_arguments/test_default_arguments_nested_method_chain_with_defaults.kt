// vybe-test: kotlin/default_arguments/test_default_arguments_nested_method_chain_with_defaults
// origin: languages/kotlin/tests/kotlin/test_default_arguments.rs

class Builder {
            fun stage(value: Int = 1): String = (value * 2).toString()
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Builder().stage()).toString(), "2")
            __check((Builder().stage(4)).toString(), "8")
        }
