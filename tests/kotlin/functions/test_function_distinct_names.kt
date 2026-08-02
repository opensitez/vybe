// vybe-test: kotlin/functions/test_function_distinct_names
// origin: languages/kotlin/tests/kotlin/test_functions.rs

fun printOne(a: Int) {
            __check((a).toString(), "10")
        }

        fun printTwo(a: Int, b: Int) {
            __check((a + b).toString(), "30")
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            printOne(10)
            printTwo(10, 20)
        }
