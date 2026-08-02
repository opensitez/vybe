// vybe-test: kotlin/type_aliases/test_typealias_generic_holder_round_trip
// origin: languages/kotlin/tests/kotlin/test_type_aliases.rs

class Holder<T>(val value: T)
        typealias StringHolder = Holder<String>

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val item: StringHolder = StringHolder("done")
            __check((item.value).toString(), "done")
        }
