// vybe-test: kotlin/local_functions/test_local_function_with_mutable_capture
// origin: languages/kotlin/tests/kotlin/test_local_functions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var count = 0
            fun inc(amount: Int) {
                count += amount
            }
            inc(2)
            inc(3)
            __check((count).toString(), "5")
        }
