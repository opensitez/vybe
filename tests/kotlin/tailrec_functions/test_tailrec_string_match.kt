// vybe-test: kotlin/tailrec_functions/test_tailrec_string_match
// origin: languages/kotlin/tests/kotlin/test_tailrec_functions.rs

tailrec fun countChars(text: String, idx: Int = 0, acc: Int = 0): Int {
            return if (idx >= text.length) acc else countChars(text, idx + 1, acc + if (text[idx] == 'a') 1 else 0)
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((countChars("abracadabra")).toString(), "4")
        }
