// vybe-test: kotlin/generics/test_generic_where_clause_with_number_and_serializable
// origin: languages/kotlin/tests/kotlin/test_generics.rs

import java.io.Serializable

        fun <T> describe(value: T): String
        where T : Number, T : Serializable {
            return value.toString()
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((describe(12)).toString(), "12")
            __check((describe(3.4)).toString(), "3.4")
        }
