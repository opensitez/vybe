// vybe-test: kotlin/named_arguments/test_named_arguments_method_with_receiver_and_defaults
// origin: languages/kotlin/tests/kotlin/test_named_arguments.rs

fun String.pad(pre: String = "<", post: String = ">"): String {
            return pre + this + post
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(("x".pad(pre = "[", post = "]")).toString(), "[x]")
            __check(("y".pad(post = "]")).toString(), "<y>")
        }
