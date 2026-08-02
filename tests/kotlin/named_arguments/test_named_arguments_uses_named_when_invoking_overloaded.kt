// vybe-test: kotlin/named_arguments/test_named_arguments_uses_named_when_invoking_overloaded
// origin: languages/kotlin/tests/kotlin/test_named_arguments.rs

fun parse(value: String): String = "s" + value
        fun parse(value: Int): String = "i" + value
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((parse(value = "x")).toString(), "sx")
            __check((parse(value = 7)).toString(), "i7")
        }
