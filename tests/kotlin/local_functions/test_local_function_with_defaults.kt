// vybe-test: kotlin/local_functions/test_local_function_with_defaults
// origin: languages/kotlin/tests/kotlin/test_local_functions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            fun greet(name: String, suffix: String = "!"): String = "hi " + name + suffix
            __check((greet("kotlin")).toString(), "hi kotlin!")
            __check((greet("kotlin", "!!")).toString(), "hi kotlin!!")
        }
