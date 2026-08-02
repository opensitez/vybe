// vybe-test: kotlin/named_arguments/test_named_arguments_mixed_with_positional
// origin: languages/kotlin/tests/kotlin/test_named_arguments.rs

fun make(a: Int, b: Int, c: Int): Int = a + b + c
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((make(1, c = 3, b = 2)).toString(), "6")
        }
