// vybe-test: kotlin/extension_functions/test_extension_function_on_function_type_calls_receiver_twice
// origin: languages/kotlin/tests/kotlin/test_extension_functions.rs

fun (() -> Int).callTwice(): Int {
            return this() + this()
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var state = 0
            val value = { state += 1
state }
            __check((value.callTwice()).toString(), "3")
            __check((value.callTwice()).toString(), "7")
            __check((state).toString(), "4")
        }
