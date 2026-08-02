// vybe-test: kotlin/functions/test_named_arguments
// origin: languages/kotlin/tests/kotlin/test_functions.rs

fun formatName(first: String, last: String): String {
            return last + ", " + first
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((formatName(last = "Smith", first = "Alice")).toString(), "Smith, Alice")
        }
