// vybe-test: kotlin/generics/test_generic_receiver_extension_with_secondary_projection
// origin: languages/kotlin/tests/kotlin/test_generics.rs

class Holder<T>(private val value: T) {
            fun value(): T = value
        }

        fun <T> Holder<T>.bind(other: T): String {
            return this.value().toString() + ":" + other.toString()
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val text = Holder("a")
            val numbers = Holder(4)
            __check((text.bind("x")).toString(), "a:x")
            __check((numbers.bind(6)).toString(), "4:6")
        }
