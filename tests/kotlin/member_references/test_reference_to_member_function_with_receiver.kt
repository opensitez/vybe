// vybe-test: kotlin/member_references/test_reference_to_member_function_with_receiver
// origin: languages/kotlin/tests/kotlin/test_member_references.rs

class Word(val value: String) {
            fun upper(): String = value.uppercase()
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val op = Word::upper
            __check((op(Word("abc"))).toString(), "ABC")
        }
