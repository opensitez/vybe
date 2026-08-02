// vybe-test: kotlin/inline_functions/test_inline_higher_order_transform
// origin: languages/kotlin/tests/kotlin/test_inline_functions.rs

inline fun <T, R> mapOrNull(value: T, transform: (T) -> R?): R? = transform(value)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((mapOrNull(4) { if (it > 2) it * 2 else null }).toString(), "8")
        }
