// vybe-test: kotlin/smart_casts/test_when_with_is_no_argument_evaluates_first_true
// origin: languages/kotlin/tests/kotlin/test_smart_casts.rs

class Dog
        class Cat

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: Any = Cat()
            val result = when {
                value is Dog -> "dog"
                value is Cat -> "cat"
                else -> "unknown"
            }
            __check((result).toString(), "cat")
        }
