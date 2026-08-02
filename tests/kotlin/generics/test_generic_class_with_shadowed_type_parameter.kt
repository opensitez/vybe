// vybe-test: kotlin/generics/test_generic_class_with_shadowed_type_parameter
// origin: languages/kotlin/tests/kotlin/test_generics.rs

class Holder<T>(val value: T) {
            fun <R> map(transform: (T) -> R): Holder<R> {
                return Holder(transform(value))
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val holder = Holder("7")
            val number = holder.map { it.toInt() }
            val text = holder.map { it + it }
            __check((number.value + 1).toString(), "8")
            __check((text.value).toString(), "77")
        }
