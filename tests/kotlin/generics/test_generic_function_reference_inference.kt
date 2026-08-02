// vybe-test: kotlin/generics/test_generic_function_reference_inference
// origin: languages/kotlin/tests/kotlin/test_generics.rs

fun <T> apply(value: T, op: (T) -> T): T {
            return op(value)
        }

        fun inc(value: Int): Int = value + 1
        fun shout(value: String): String = value.toUpperCase()

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((apply(2, ::inc)).toString(), "3")
            __check((apply("ok", ::shout)).toString(), "OK")
        }
