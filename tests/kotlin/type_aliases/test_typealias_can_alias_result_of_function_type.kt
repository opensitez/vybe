// vybe-test: kotlin/type_aliases/test_typealias_can_alias_result_of_function_type
// origin: languages/kotlin/tests/kotlin/test_type_aliases.rs

typealias Next = () -> Int

        fun make(value: Int): Next {
            return { value + 1 }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val next: Next = make(7)
            __check((next()).toString(), "8")
        }
