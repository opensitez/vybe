// vybe-test: kotlin/spread_arguments/test_spread_in_nested_calls
// origin: languages/kotlin/tests/kotlin/test_spread_arguments.rs

fun add(a: Int, b: Int, c: Int): Int = a + b + c
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val head = intArrayOf(1, 2)
            __check((add(3, *head)).toString(), "6")
        }
