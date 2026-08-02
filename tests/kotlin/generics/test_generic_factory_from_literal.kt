// vybe-test: kotlin/generics/test_generic_factory_from_literal
// origin: languages/kotlin/tests/kotlin/test_generics.rs

fun <T> one(value: T): Array<T> {
            return arrayOf(value)
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val numbers = one(42)
            val words = one("zap")
            __check((numbers.size).toString(), "1")
            __check((numbers[0]).toString(), "42")
            __check((words[0]).toString(), "zap")
        }
