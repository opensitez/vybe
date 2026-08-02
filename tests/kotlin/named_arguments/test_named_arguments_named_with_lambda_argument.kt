// vybe-test: kotlin/named_arguments/test_named_arguments_named_with_lambda_argument
// origin: languages/kotlin/tests/kotlin/test_named_arguments.rs

fun map(value: Int, transform: (Int) -> Int = { it }, offset: Int = 0): Int {
            return transform(value) + offset
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((map(value = 2, offset = 1)).toString(), "3")
            __check((map(3, transform = { it * it })).toString(), "9")
        }
