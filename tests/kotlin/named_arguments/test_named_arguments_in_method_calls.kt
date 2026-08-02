// vybe-test: kotlin/named_arguments/test_named_arguments_in_method_calls
// origin: languages/kotlin/tests/kotlin/test_named_arguments.rs

class Tagger {
            fun compose(head: String, value: String, tail: String): String = head + value + tail
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val t = Tagger()
            __check((t.compose(head = "[", value = "v", tail = "]")).toString(), "[v]")
        }
