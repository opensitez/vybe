// vybe-test: kotlin/named_arguments/test_named_arguments_with_nullable_receiver
// origin: languages/kotlin/tests/kotlin/test_named_arguments.rs

fun choose(first: String?, second: String = "b"): String {
            return (first ?: second)
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((choose(first = null, second = "k")).toString(), "k")
        }
