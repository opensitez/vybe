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

        var __buf: String = ""

fun __p(s: String) {
    __buf = __buf + s + "\n"
}

fun __pr(s: String) {
    __buf = __buf + s
}

// The final `println` contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted. Written as two equality
// tests rather than trimming: `String.endsWith` is not implemented in Vybe's
// Kotlin (measured — `"ab\n".endsWith("\n")` throws "undefined is not
// callable"), and a harness that cannot run asserts nothing at all. The cargo
// helper split on "\n" and popped trailing empties, so the two forms were
// equivalent there too.
fun __check(want: String) {
    if (__buf != want && __buf != want + "\n") {
        println("FAIL: want [" + want + "] got [" + __buf + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val converter: Converter<Any, Number> = StringToInt()
            __p((emit(converter, "abc")).toString())
        
__check("3")
}
