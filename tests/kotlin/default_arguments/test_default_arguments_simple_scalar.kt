// vybe-test: kotlin/default_arguments/test_default_arguments_simple_scalar
// origin: languages/kotlin/tests/kotlin/test_default_arguments.rs

fun greet(name: String = "world"): String = "hi " + name
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((greet()).toString(), "hi world")
            __check((greet("kotlin")).toString(), "hi kotlin")
        }
