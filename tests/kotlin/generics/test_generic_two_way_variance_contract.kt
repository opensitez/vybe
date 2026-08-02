// vybe-test: kotlin/generics/test_generic_two_way_variance_contract
// origin: languages/kotlin/tests/kotlin/test_generics.rs

interface Converter<in S, out T> {
            fun convert(value: S): T
        }

        class StringToInt : Converter<String, Number> {
            override fun convert(value: String): Number = value.length
        }

        fun emit(any: Converter<CharSequence, Number>, value: CharSequence): String {
            return any.convert(value).toString()
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val converter: Converter<Any, Number> = StringToInt()
            __check((emit(converter, "abc")).toString(), "3")
        }
