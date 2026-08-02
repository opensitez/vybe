// vybe-test: kotlin/generics/test_generic_function_returning_array_and_size
// origin: languages/kotlin/tests/kotlin/test_generics.rs

fun <T> toArray(left: T, right: T): Array<T> {
            return arrayOf(left, right)
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val numbers = toArray(2, 3)
            val words = toArray("a", "b")
            __check((numbers.size).toString(), "2")
            __check((words.size).toString(), "2")
            __check((numbers[1] + words[1]).toString(), "3b")
        }
