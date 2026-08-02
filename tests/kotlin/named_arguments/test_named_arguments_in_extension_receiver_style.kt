// vybe-test: kotlin/named_arguments/test_named_arguments_in_extension_receiver_style
// origin: languages/kotlin/tests/kotlin/test_named_arguments.rs

fun Int.scale(base: Int = 2, times: Int = 1): Int {
            return this * base + times
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((3.scale(times = 5)).toString(), "11")
            __check((3.scale(base = 4, times = 1)).toString(), "13")
        }
