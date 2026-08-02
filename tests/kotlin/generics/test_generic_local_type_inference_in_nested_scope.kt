// vybe-test: kotlin/generics/test_generic_local_type_inference_in_nested_scope
// origin: languages/kotlin/tests/kotlin/test_generics.rs

class Holder<T>(private val value: T) {
            fun value(): T = value
        }

        fun <T> describe(value: Holder<T>): String {
            return value.value().toString()
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val number = Holder(8)
            val text = Holder("gen")
            val inferred = describe(number)
            val direct = describe(text)
            __check((inferred).toString(), "8")
            __check((direct).toString(), "gen")
        }
