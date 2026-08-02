// vybe-test: kotlin/functions/test_function_block_body
// origin: languages/kotlin/tests/kotlin/test_functions.rs

fun greet(name: String) {
            __check(("Hi " + name).toString(), "Hi Kotlin")
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            greet("Kotlin")
        }
