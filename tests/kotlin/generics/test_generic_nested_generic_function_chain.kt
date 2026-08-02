// vybe-test: kotlin/generics/test_generic_nested_generic_function_chain
// origin: languages/kotlin/tests/kotlin/test_generics.rs

fun <T> wrap(value: T): Box<T> = Box(value)
        class Box<T>(val value: T)

        fun <T, R> wrapChain(value: T, op: (T) -> R): Box<R> {
            return wrap(op(value))
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val boxed = wrapChain(3, { it + 1 })
            val text = wrapChain("ok", { it + "!" })
            __check((boxed.value).toString(), "4")
            __check((text.value).toString(), "ok!")
        }
