// vybe-test: kotlin/named_arguments/test_named_arguments_named_and_default_in_chained_calls
// origin: languages/kotlin/tests/kotlin/test_named_arguments.rs

fun base(x: Int = 1, y: Int = 2): Int = x + y
        fun scale(x: Int, y: Int = 2): Int = x * y
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((base(y = 9)).toString(), "10")
            __check((scale(3, y = 5)).toString(), "15")
        }
