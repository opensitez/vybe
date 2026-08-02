// vybe-test: kotlin/functions/test_function_string_return
// origin: languages/kotlin/tests/kotlin/test_functions.rs

fun concat(a: String, b: String): String = a + b

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((concat("foo", "bar")).toString(), "foobar")
        }
