// vybe-test: kotlin/variance/test_variance_nested_generics_covariant_nested
// origin: languages/kotlin/tests/kotlin/test_variance.rs

interface Boxed<out T> { val value: T }
        class Holder<T>(override val value: T) : Boxed<T>
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val boxed: Boxed<String> = Holder("k")
            val anyBox: Boxed<Any> = boxed
            __check((anyBox.value).toString(), "k")
        }
