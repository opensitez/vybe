// vybe-test: kotlin/generics/test_generic_class_with_method
// origin: languages/kotlin/tests/kotlin/test_generics.rs

class Cache<T>(initial: T) {
            private val value: T = initial
            fun unwrap(): T {
                return value
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Cache("hello").unwrap()).toString(), "hello")
            __check((Cache(8).unwrap()).toString(), "8")
        }
