// vybe-test: kotlin/local_functions/test_local_function_without_parameters
// origin: languages/kotlin/tests/kotlin/test_local_functions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var value = 0
            fun tick() { value += 1 }
            tick()
            tick()
            __check((value).toString(), "2")
        }
